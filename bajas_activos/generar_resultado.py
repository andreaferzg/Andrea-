from __future__ import annotations

from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import posixpath
import re
from typing import Any
import unicodedata
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from rapidfuzz import fuzz, process


ROOT = Path(__file__).resolve().parent
BASE_FOLDER = ROOT / "BASE DESCARGADA DEL PORTAL"
EXPORT_FOLDER = ROOT / "IH08"
AR02_FOLDER = ROOT / "AR02"
# AR02 de R3 (antes de la migración a HANA). Opcional: si la carpeta no
# existe o está vacía, el cruce con el AR02 antiguo simplemente se omite.
AR02_LEGACY_FOLDER = ROOT / "AR02 ANTIGUO"
# Si el activo HANA está con todos sus valores en 0 (migración R3 → HANA mal
# hecha) y el AR02 ANTIGUO sí tiene valor, ¿usar el valor antiguo como VALOR
# CONTABLE del STATUS? En False solo se avisa en STATUS - REVISAR.
USE_LEGACY_BOOK_VALUE_WHEN_HANA_IS_ZERO = False
# Encabezados (normalizados con normalize_header) que se reconocen como
# "placa antigua" en IH08 y como "N.º de activo antiguo (R3)" en el AR02 nuevo.
IH08_LEGACY_PLATE_HEADERS = {
    "placa antigua",
    "placa anterior",
    "placa r3",
    "placa antiguo",
    "codigo af r3",
    "af antiguo",
    "activo antiguo",
    "activo fijo antiguo",
}
AR02_LEGACY_ASSET_HEADERS = {
    "activo fijo orig",
    "activo fijo original",
    "af orig",
    "activo fijo antiguo",
    "activo antiguo",
    "af antiguo",
    "n af antiguo",
    "nro af antiguo",
    "numero af antiguo",
    "numero de af antiguo",
    "numero activo antiguo",
    "numero de activo antiguo",
    "n activo antiguo",
    "codigo af r3",
    "activo r3",
    "af r3",
    "activo fijo r3",
    "n activo fijo antiguo",
    "numero de activo fijo antiguo",
}
ACDOCA_FOLDER = ROOT / "TABLA ACDOCA"
BKPF_FOLDER = ROOT / "TABLA BKPF"
OUTPUT_FOLDER = ROOT / "RESULTADO COMPLETO"
STATUS_OUTPUT_FOLDER = ROOT / "RESULTADOS STATUS"
STATUS_TEMPLATE_PATH = ROOT / "STATUS DE BAJA" / "00 STATUS DE BAJA 2026.xlsx"
STATUS_SHEET_NAME = "DATA FINAL 2026"
# Clases de movimiento de BAJA en ACDOCA (en AG10 las bajas salen como Z13/Z11,
# además de las estándar 2xx). La 980 es la carga de migración: no es baja.
BAJA_MOVEMENT_PATTERN = re.compile(r"2\d{2}|Z11|Z13", re.IGNORECASE)
BASE_SHEET_MARKER = "CONSOLIDADO"
BASE_HEADER_ROW = 6
BASE_EQUIPMENT_COLUMN = 19
BASE_MATERIAL_COLUMN = 20
BASE_SERIAL_COLUMN = 21
EXPORT_EQUIPMENT_COLUMN = 3
EXPORT_MATERIAL_COLUMN = 12
EXPORT_SERIAL_COLUMN = 18
ACDOCA_OUTPUT_HEADERS = [
    "N.º DOCUMENTO CONTABLE SAP",
    "CUENTA CONTABLE SAP",
    "IMPORTE CONTABLE SAP",
    "FECHA CONTABILIZACION SAP",
    "ESTADO CONCILIACIÓN ACDOCA",
    "FILA BASE ORIGEN",
    "FILA ACDOCA ORIGEN",
]
AR02_OUTPUT_HEADERS = [
    "ESTADO VALIDACIÓN AR02",
    "ESTADO MOVIMIENTOS ACDOCA",
    "CLAVE ACTIVO AR02",
    "DENOMINACIÓN AR02",
    "CUENTA BALANCE CAP AR02",
    "CLASE ACTIVOS AR02",
    "CAPITALIZADO EL AR02",
    "CAP INIEJ AR02",
    "ALTA AR02",
    "ANTICIPO AR02",
    "BAJA AR02",
    "AMO BAJA AR02",
    "TRASLADO AR02",
    "CAP ACTUALES AR02",
    "AMO ACUMULADA AR02",
    "VALCONT INIEJ AR02",
    "VALCONT ACTUAL AR02",
    "ESTADO ACTIVO AR02",
    "DOCUMENTOS ACDOCA DEL ACTIVO",
    "N.º DOCUMENTOS ACDOCA DEL ACTIVO",
    "CUENTAS 32/33 ACDOCA",
    "CONTROL CUENTA AR02",
    "FILA AR02 ORIGEN",
    "RUTA VALIDACIÓN AR02",
    "PLACA ANTIGUA IH08",
    "AF R3 (ANTIGUO)",
    "ORIGEN AF R3",
    "ESTADO AR02 ANTIGUO",
    "DENOMINACIÓN AR02 ANTIGUO",
    "CAP ACTUALES AR02 ANTIGUO",
    "BAJA AR02 ANTIGUO",
    "AMO ACUMULADA AR02 ANTIGUO",
    "VALCONT ACTUAL AR02 ANTIGUO",
    "ESTADO ACTIVO AR02 ANTIGUO",
    "COMPARACIÓN AR02 HANA VS ANTIGUO",
    "FILA AR02 ANTIGUO ORIGEN",
]
# Columnas opcionales del AR02 HANA con placas. ('Equipo Antiguo' NO se usa:
# es la numeración de equipos de R3 y choca con los equipos actuales de IH08.)
AR02_PLATE_HEADERS = ("Placa Activo Fijo", "Placa antigua")
# Columnas del AR02 ANTIGUO (R3) que se leen; se toma la primera aparición
# porque el archivo trae a la derecha otro bloque con nombres repetidos.
AR02_LEGACY_HEADERS = {
    "anln1": "Activo fijo",
    "anln2": "Subnúmero",
    "denom": "Denominación del activo fijo",
    "cap": "CAP actuales",
    "baja": "Baja",
    "amo_baja": "Amo baja",
    "amo": "Amo acumulada",
    "valcont": "Val.cont.act.",
}
# Mismo orden que la hoja DATA FINAL 2026 (columnas A:AJ).
STATUS_HEADERS = [
    "DOC",
    "MES TRABAJADO",
    "N°",
    "EMISOR",
    "DOC. REFERENCIA",
    "MES RECEPCION FICHA DE ESTADO",
    "FECHA DE RECEPCION",
    "ORD CAPEX SAP",
    "CODIGO AF R3",
    "CODIGO AF HANAH",
    "VALOR CONTABLE",
    "DESCRIPCION DEL ACTIVO",
    "PLACA",
    "MARCA",
    "MODELO",
    "Nº SERIE",
    "SERIE SAP",
    "NUEVO ESTADO",
    "xxx",
    "CENTRO",
    "E/P",
    "SE DIO DE BAJA",
    "MES DE BAJA",
    "N° DE DOCUMENTO",
    "COMENTARIOS",
    "TIPO I",
    "TIPO II",
    "AUXILIAR  1",
    "AUXILIAR 2",
    "AUXILIAR 3",
    "AUXILIAR 4",
    "COMENTARIOS2",
    "AUXILIAR 6",
    "COORDINACION SOBRE BAJAS",
    "N° FACTURA VENTAS",
    "FECHA",
]
STATUS_TRACE_HEADERS = [
    "FILA STATUS",
    "TIPO DE FILA",
    "SERIE SAP",
    "COLUMNA",
    "VALOR ESCRITO",
    "FUENTE / RUTA",
]
STATUS_REVIEW_HEADERS = [
    "FILA STATUS",
    "SERIE SAP",
    "CODIGO AF HANAH (STATUS)",
    "ALERTA",
    "DETALLE",
]
SPANISH_MONTHS = (
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SETIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
)
NAVY = "17365D"
BLUE = "28527A"
TEAL = "167D8D"
GOLD = "A66B00"
TEXT = "243247"
MUTED = "64748B"

ACDOCA_REQUIRED_HEADERS = {
    "rldnr": {"ledger", "rldnr"},
    "rbukrs": {"sociedad", "rbukrs"},
    "gjahr": {"ejercicio", "gjahr"},
    "belnr": {"n documento", "no documento", "numero documento", "belnr"},
    "racct": {"numero de cuenta", "racct"},
    "hsl": {"importe en moneda de sociedad", "hsl"},
    "budat": {"fecha contabiliz", "fecha contabilizacion", "budat"},
    "anln1": {"activo fijo", "anln1"},
    "anln2": {"subnumero", "anln2"},
    "docln": {"pos doc libro mayor", "docln"},
    "anbwa": {"clase movimiento activos fijos", "anbwa"},
}
ACDOCA_OPTIONAL_HEADERS = {
    "blart": {"clase de documento", "blart"},
}
AR02_REQUIRED_HEADERS = {
    "anln1": {"activo fijo", "anln1"},
    "anln2": {"subnumero", "sn"},
    "denom": {"denominacion del activo fijo", "denominacion", "txt50"},
    "cta_cap": {"cuenta balance cap", "cta cap"},
    "clase": {"clase activos fijos", "clase", "anlkl"},
    "capitalized": {"capitalizado el"},
    "center": {"centro"},
}
AR02_VALUE_HEADERS = (
    "CAP iniEj",
    "Alta",
    "Anticipo",
    "Baja",
    "Amo baja",
    "Traslado",
    "CAP actuales",
    "Amo acumulada",
    "ValCont. iniEj",
    "Val.cont.act.",
)
# Columnas de valores que pueden no venir en el layout del AR02 (quedan vacías).
AR02_OPTIONAL_VALUE_HEADERS = {"Anticipo"}
BKPF_REQUIRED_HEADERS = {
    "bukrs": {"sociedad", "bukrs"},
    "belnr": {"n documento", "no documento", "numero documento", "belnr"},
    "gjahr": {"ejercicio", "gjahr"},
    "budat": {"fecha contabiliz", "fecha contabilizacion", "budat"},
    "awtyp": {"operacion referencia", "awtyp"},
    "awkey": {"clave de referencia", "awkey"},
}
BKPF_PART_PATTERN = re.compile(
    r"^(?P<base>.*?)[ _.-]+(?:part|parte|pt)[ _.-]*(?P<number>\d+)"
    r"(?:[ _.-]+(?:of|de)[ _.-]*(?P<total>\d+))?$",
    re.IGNORECASE,
)


def newest_excel(folder: Path) -> Path:
    if not folder.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {folder}")

    files = [
        path
        for path in folder.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"No se encontraron archivos .xlsx en: {folder}")

    return max(files, key=lambda path: path.stat().st_mtime)


def newest_excel_parts(folder: Path) -> list[Path]:
    if not folder.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {folder}")

    files = [
        path
        for path in folder.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"No se encontraron archivos .xlsx en: {folder}")

    newest = max(files, key=lambda path: path.stat().st_mtime)
    match = BKPF_PART_PATTERN.match(newest.stem)
    if match is None:
        return [newest]

    base_name = match.group("base").casefold()
    parts: dict[int, Path] = {}
    declared_total = int(match.group("total")) if match.group("total") else None
    for path in files:
        part_match = BKPF_PART_PATTERN.match(path.stem)
        if part_match is None or part_match.group("base").casefold() != base_name:
            continue
        part_number = int(part_match.group("number"))
        if part_number in parts:
            raise ValueError(
                f"Hay más de un archivo para la parte {part_number} de BKPF "
                f"({base_name})."
            )
        parts[part_number] = path
        part_total = (
            int(part_match.group("total")) if part_match.group("total") else None
        )
        if part_total is not None:
            if declared_total is not None and declared_total != part_total:
                raise ValueError(
                    f"Los archivos divididos de BKPF declaran totales incompatibles "
                    f"({declared_total} y {part_total})."
                )
            declared_total = part_total

    expected_parts = (
        set(range(1, declared_total + 1))
        if declared_total is not None
        else set(range(1, max(parts) + 1))
    )
    if set(parts) != expected_parts:
        raise ValueError(
            "La exportación BKPF parece estar dividida, pero faltan partes o "
            f"la numeración no es continua. Encontradas: {sorted(parts)}."
        )
    return [parts[number] for number in sorted(parts)]


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_key(value: Any) -> str:
    return cell_text(value).casefold()


def normalize_numeric_identifier(value: Any, width: int) -> str:
    text = cell_text(value)
    if not text or not text.isalnum() or len(text) > width:
        return ""
    return text.zfill(width) if text.isdigit() else text.upper()


def normalize_asset_code(value: Any) -> str:
    text = cell_text(value).removesuffix(".0").strip()
    return text.zfill(12) if text.isdigit() else ""


def normalize_asset_subnumber(value: Any) -> str:
    text = cell_text(value).removesuffix(".0").strip()
    return text.zfill(4) if text.isdigit() else ""


def is_current_ih08_row(value: Any) -> bool:
    if isinstance(value, datetime):
        return value.date() == datetime(9999, 12, 31).date()
    text = cell_text(value).strip()
    return bool(
        re.fullmatch(
            r"(?:31[./-]12[./-]9999|9999[./-]12[./-]31)(?:\s+00:00:00)?",
            text,
        )
    )


def normalize_account(value: Any) -> str:
    text = cell_text(value).removesuffix(".0").strip()
    return text if text.isdigit() else text


def normalize_status_plate(value: Any) -> str:
    plate = re.sub(r"\s+", "", cell_text(value)).upper()
    if not plate:
        return ""
    if re.fullmatch(r"P0\d-?\d{6}", plate):
        return f"{plate[:3]}-{plate[-6:]}"
    if plate.startswith(("A", "F")) and plate[1:].isdigit():
        width = 11 if plate.startswith("A") else 9
        if len(plate[1:]) <= width:
            return plate[0] + plate[1:].zfill(width)
    return plate


def excel_serial(value: datetime) -> int:
    return (value.date() - datetime(1899, 12, 30).date()).days


def is_plate_like(value: Any) -> bool:
    plate = normalize_status_plate(value)
    return bool(
        re.fullmatch(r"[PF]\d{2}-\d{6}", plate)
        or re.fullmatch(r"A\d{11}", plate)
    )


def fic_number(value: Any) -> str:
    """Devuelve el número de ficha sin ceros ('FIC-000125' -> '125')."""
    text = cell_text(value)
    if text.isdigit():
        return str(int(text))
    match = re.search(r"FIC\D{0,3}0*(\d+)", text, re.IGNORECASE)
    return match.group(1) if match else ""


def status_amount(value: Any) -> Decimal:
    try:
        amount = Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as error:
        raise ValueError(
            f"VALOR CONTABLE no es numérico: {value!r}."
        ) from error
    if not amount.is_finite():
        raise ValueError(f"VALOR CONTABLE no es finito: {value!r}.")
    return amount


def normalize_search_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", cell_text(value))
    text = "".join(
        character for character in text if not unicodedata.combining(character)
    )
    return re.sub(r"[^a-z0-9]+", " ", text.casefold()).strip()


def read_ar02(
    path: Path,
) -> tuple[list[str], list[tuple[int, list[Any]]], dict[str, int]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows()
        header_cells = next(rows, None)
        if header_cells is None:
            raise ValueError(f"El archivo AR02 está vacío: {path.name}.")

        raw_headers = tuple(cell.value for cell in header_cells)
        headers = unique_headers(raw_headers, prefix="Campo AR02")
        normalized_headers = [normalize_header(value) for value in raw_headers]
        column_indices: dict[str, int] = {}
        for field, aliases in AR02_REQUIRED_HEADERS.items():
            matches = [
                index
                for index, header in enumerate(normalized_headers)
                if header in aliases
            ]
            if len(matches) != 1:
                raise ValueError(
                    f"Se esperaba una columna única para AR02.{field} en "
                    f"{path.name}; se encontraron {len(matches)}."
                )
            column_indices[field] = matches[0]
        legacy_matches = find_legacy_header(
            normalized_headers,
            AR02_LEGACY_ASSET_HEADERS,
            loose_words={"activo", "af"},
        )
        if len(legacy_matches) == 1:
            column_indices["legacy"] = legacy_matches[0]
        elif len(legacy_matches) > 1:
            raise ValueError(
                "Hay más de una columna de N.º de activo antiguo (R3) en "
                f"{path.name}: "
                f"{', '.join(headers[index] for index in legacy_matches)}."
            )

        records: list[tuple[int, list[Any]]] = []
        for source_row, cells in enumerate(rows, start=2):
            row = [cell_text(cell.value) for cell in cells]
            asset_code = normalize_asset_code(row[column_indices["anln1"]])
            if not asset_code:
                continue
            row[column_indices["anln1"]] = asset_code
            row[column_indices["anln2"]] = (
                normalize_asset_subnumber(row[column_indices["anln2"]])
                or "0000"
            )
            if "cta_cap" in column_indices:
                row[column_indices["cta_cap"]] = normalize_account(
                    row[column_indices["cta_cap"]]
                )
            records.append((source_row, row))
    finally:
        workbook.close()

    if not records:
        raise ValueError(f"No se encontraron activos válidos en {path.name}.")
    return headers, records, column_indices


def find_legacy_header(
    normalized_headers: list[str],
    aliases: set[str],
    loose_words: set[str] = frozenset({"activo", "af", "placa"}),
) -> list[int]:
    """Índices de la columna 'antigua': primero coincidencia exacta con un
    alias; si no hay, encabezados que digan 'antigu'/'r3' junto a
    'activo'/'af'/'placa'."""
    exact = [
        index
        for index, header in enumerate(normalized_headers)
        if header in aliases
    ]
    if exact:
        return exact
    loose = []
    for index, header in enumerate(normalized_headers):
        words = set(header.split())
        if not words & {"antiguo", "antigua", "anterior", "r3"}:
            continue
        if words & loose_words:
            loose.append(index)
    return loose


def normalize_legacy_asset(value: Any) -> tuple[str, str]:
    """'5600001082', '5600001082-0', 'AF 5600001082' -> ('5600001082', '0000').

    Devuelve ('', '') si el valor no tiene forma de N.º de activo R3
    (exactamente 10 dígitos, con subnúmero opcional)."""
    text = cell_text(value).removesuffix(".0").strip().upper()
    text = re.sub(r"^(?:AF|ACTIVO|R3)\W*", "", text)
    match = re.fullmatch(r"(\d{10})(?:\s*[-/ ]\s*(\d{1,4}))?", text)
    if not match:
        return "", ""
    subnumber = match.group(2).zfill(4) if match.group(2) else ""
    return match.group(1).zfill(10), subnumber


def split_plates(value: Any) -> set[str]:
    """'P07-001508 / A00000010120' -> {'P07-001508', 'A00000010120'}.

    Solo devuelve textos con forma de placa (P/F + 2 dígitos + 6, o A + 11)."""
    text = cell_text(value).replace("–", "-").replace("—", "-")
    plates: set[str] = set()
    for part in re.split(r"[/;,\n]|\s+Y\s+", text, flags=re.IGNORECASE):
        plate = normalize_status_plate(re.sub(r"\s*-\s*", "-", part.strip()))
        if is_plate_like(plate):
            plates.add(plate)
    return plates


def read_ar02_legacy(
    folder: Path,
) -> tuple[Path | None, dict[tuple[str, str], dict[str, Any]]]:
    """Lee el AR02 ANTIGUO (R3). Clave: (activo R3 de 10 dígitos, subnúmero).

    Si un mismo activo aparece en varias filas, se suman los importes."""
    if not folder.is_dir():
        return None, {}
    try:
        path = newest_excel(folder)
    except FileNotFoundError:
        return None, {}
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        wanted = {
            field: normalize_header(header)
            for field, header in AR02_LEGACY_HEADERS.items()
        }
        header_row = None
        indices: dict[str, int] = {}
        for row_number, cells in enumerate(
            worksheet.iter_rows(max_row=20, values_only=True), start=1
        ):
            normalized = [normalize_header(value) for value in cells]
            found: dict[str, int] = {}
            for field, header in wanted.items():
                if header in normalized:
                    found[field] = normalized.index(header)
            if len(found) == len(wanted):
                header_row, indices = row_number, found
                break
        if header_row is None:
            missing = ", ".join(AR02_LEGACY_HEADERS.values())
            raise ValueError(
                f"No se encontró la fila de encabezados del AR02 ANTIGUO en "
                f"{path.name} (se buscan: {missing})."
            )

        legacy: dict[tuple[str, str], dict[str, Any]] = {}
        for source_row, cells in enumerate(
            worksheet.iter_rows(min_row=header_row + 1, values_only=True),
            start=header_row + 1,
        ):
            if len(cells) <= max(indices.values()):
                continue
            code, _ = normalize_legacy_asset(cells[indices["anln1"]])
            if not code:
                continue
            subnumber = normalize_asset_subnumber(cells[indices["anln2"]]) or "0000"
            entry = legacy.setdefault(
                (code, subnumber),
                {
                    "denom": cell_text(cells[indices["denom"]]),
                    "rows": [],
                    **{field: Decimal(0) for field in ("cap", "baja", "amo_baja", "amo", "valcont")},
                },
            )
            entry["rows"].append(source_row)
            for field in ("cap", "baja", "amo_baja", "amo", "valcont"):
                value = cells[indices[field]]
                if isinstance(value, str) and value.startswith("="):
                    continue  # fórmula sin valor guardado
                try:
                    entry[field] += amount_key(value) or Decimal(0)
                except ValueError:
                    continue
    finally:
        workbook.close()
    return path, legacy


def legacy_asset_state(entry: dict[str, Any]) -> str:
    if entry["cap"] == 0 and (entry["baja"] != 0 or entry["amo_baja"] != 0):
        return "BAJA EN AR02 ANTIGUO"
    if entry["cap"] != 0 and entry["valcont"] == 0:
        return "TOTALMENTE DEPRECIADO"
    return ""


def date_key(value: Any) -> Any:
    return value.date() if isinstance(value, datetime) else value


def amount_key(value: Any) -> Decimal | None:
    if value is None or not cell_text(value):
        return None
    try:
        return Decimal(str(value))
    except InvalidOperation as error:
        raise ValueError(f"Importe contable inválido: {value!r}.") from error


def ar02_asset_state(
    record: list[Any],
    column_indices: dict[str, int],
) -> str:
    cap_current = amount_key(
        record[column_indices[normalize_header("CAP actuales")]]
    )
    disposal = (
        amount_key(record[column_indices[normalize_header("Baja")]])
        or Decimal(0)
    ) != 0 or (
        amount_key(record[column_indices[normalize_header("Amo baja")]])
        or Decimal(0)
    ) != 0
    if cap_current == 0 and disposal:
        return "BAJA EN AR02"
    return ""


def identifier_value(value: Any, number_format: str = "General") -> Any:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if re.fullmatch(r"0+", number_format):
        return cell_text(value).zfill(len(number_format))
    return cell_text(value)


def normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", cell_text(value))
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = text.replace("°", " ").replace("º", " ").casefold()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def read_acdoca(
    path: Path,
) -> tuple[list[str], list[tuple[int, list[Any]]], dict[str, int]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = worksheet.iter_rows()
        header_cells = next(rows, None)
        if header_cells is None:
            raise ValueError(f"El archivo ACDOCA está vacío: {path.name}.")

        raw_headers = tuple(cell.value for cell in header_cells)
        headers = unique_headers(raw_headers, prefix="Campo ACDOCA")
        normalized_headers = [normalize_header(value) for value in raw_headers]
        column_indices: dict[str, int] = {}
        for field, aliases in ACDOCA_REQUIRED_HEADERS.items():
            matches = [
                index
                for index, header in enumerate(normalized_headers)
                if header in aliases
            ]
            if len(matches) != 1:
                raise ValueError(
                    f"Se esperaba una columna única para {field} en {path.name}; "
                    f"se encontraron {len(matches)}."
                )
            column_indices[field] = matches[0]

        for field, aliases in ACDOCA_OPTIONAL_HEADERS.items():
            matches = [
                index
                for index, header in enumerate(normalized_headers)
                if header in aliases
            ]
            if len(matches) > 1:
                raise ValueError(
                    f"Se esperaba una columna única para ACDOCA.{field} en "
                    f"{path.name}; se encontraron {len(matches)}."
                )
            if matches:
                column_indices[field] = matches[0]

        identifier_fields = {"belnr", "racct", "anln1", "anln2"}
        records: list[tuple[int, list[Any]]] = []
        seen_positions: set[tuple[str, str, str, str, str]] = set()
        for source_row, cells in enumerate(rows, start=2):
            row = [cell.value for cell in cells]
            for field in identifier_fields:
                index = column_indices[field]
                row[index] = identifier_value(
                    row[index],
                    cells[index].number_format,
                )
            row[column_indices["racct"]] = normalize_account(
                row[column_indices["racct"]]
            )
            row[column_indices["anln1"]] = normalize_asset_code(
                row[column_indices["anln1"]]
            )
            row[column_indices["anln2"]] = (
                normalize_asset_subnumber(row[column_indices["anln2"]])
                or "0000"
            )
            if not any(value is not None and str(value).strip() for value in row):
                continue
            if (
                normalize_key(row[column_indices["rbukrs"]]) != "ag10"
                or normalize_key(row[column_indices["rldnr"]]) != "0l"
                or not normalize_key(row[column_indices["racct"]]).startswith(
                    ("32", "33")
                )
            ):
                continue
            required_values = ("belnr", "racct", "docln", "budat")
            if any(
                not cell_text(row[column_indices[field]])
                for field in required_values
            ):
                raise ValueError(
                    f"Hay una posición ACDOCA incompleta en la fila {source_row} "
                    f"de {path.name}."
                )
            position_key = (
                normalize_key(row[column_indices["rldnr"]]),
                normalize_key(row[column_indices["rbukrs"]]),
                normalize_key(row[column_indices["gjahr"]]),
                normalize_numeric_identifier(row[column_indices["belnr"]], 10),
                normalize_key(row[column_indices["docln"]]),
            )
            if position_key in seen_positions:
                raise ValueError(
                    f"Posición ACDOCA duplicada en la fila {source_row} "
                    f"de {path.name}: {position_key!r}."
                )
            seen_positions.add(position_key)
            records.append((source_row, list(row)))
    finally:
        workbook.close()

    return headers, records, column_indices


def read_bkpf(
    paths: list[Path],
) -> tuple[
    list[str],
    list[tuple[str, int, list[Any]]],
    dict[str, int],
]:
    records: list[tuple[str, int, list[Any]]] = []
    column_indices: dict[str, int] | None = None
    expected_headers: list[str] | None = None
    seen_documents: set[tuple[str, str, str]] = set()

    for path in paths:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            worksheet = workbook.active
            rows = worksheet.iter_rows()
            header_cells = next(rows, None)
            if header_cells is None:
                raise ValueError(f"El archivo BKPF está vacío: {path.name}.")

            raw_headers = tuple(cell.value for cell in header_cells)
            normalized_headers = [normalize_header(value) for value in raw_headers]
            headers = unique_headers(raw_headers, prefix="Campo BKPF")
            current_indices: dict[str, int] = {}
            for field, aliases in BKPF_REQUIRED_HEADERS.items():
                matches = [
                    index
                    for index, header in enumerate(normalized_headers)
                    if header in aliases
                ]
                if len(matches) != 1:
                    raise ValueError(
                        f"Se esperaba una columna única para BKPF.{field} en "
                        f"{path.name}; se encontraron {len(matches)}."
                    )
                current_indices[field] = matches[0]

            if expected_headers is None:
                expected_headers = headers
                column_indices = current_indices
            elif headers != expected_headers:
                raise ValueError(
                    "Las partes de BKPF tienen estructuras incompatibles. "
                    f"Revisa {path.name}."
                )

            for source_row, cells in enumerate(rows, start=2):
                row = [cell.value for cell in cells]
                if not any(value is not None and str(value).strip() for value in row):
                    continue
                for field in ("bukrs", "belnr", "gjahr", "awtyp", "awkey", "budat"):
                    index = current_indices[field]
                    if not cell_text(row[index]):
                        raise ValueError(
                            f"BKPF.{field} está vacío en {path.name}, "
                            f"fila {source_row}."
                        )

                company = normalize_key(row[current_indices["bukrs"]])
                year = normalize_key(row[current_indices["gjahr"]])
                belnr = normalize_numeric_identifier(
                    row[current_indices["belnr"]], 10
                )
                if company != "ag10" or year != "2026" or not belnr:
                    raise ValueError(
                        f"El alcance de BKPF no es compatible en {path.name}, "
                        f"fila {source_row}: sociedad={company!r}, "
                        f"ejercicio={year!r}, documento="
                        f"{cell_text(row[current_indices['belnr']])!r}."
                    )
                row[current_indices["belnr"]] = belnr
                document_key = (company, year, belnr)
                if document_key in seen_documents:
                    raise ValueError(
                        "Documento BKPF duplicado entre archivos/filas: "
                        f"{document_key!r} ({path.name}, fila {source_row})."
                    )
                seen_documents.add(document_key)
                records.append((path.name, source_row, row))
        finally:
            workbook.close()

    if not records or expected_headers is None or column_indices is None:
        raise ValueError("No se encontraron registros BKPF válidos.")
    return expected_headers, records, column_indices


def populated_column_indices(
    headers: list[str], records: list[list[Any]]
) -> list[int]:
    return [
        index
        for index in range(len(headers))
        if any(
            index < len(record)
            and record[index] is not None
            and str(record[index]).strip()
            for record in records
        )
    ]


def unique_headers(values: tuple[Any, ...], prefix: str = "Columna") -> list[str]:
    headers: list[str] = []
    counts: Counter[str] = Counter()
    for index, value in enumerate(values, start=1):
        label = cell_text(value) or f"{prefix} {index}"
        counts[label] += 1
        if counts[label] > 1:
            label = f"{label} ({counts[label]})"
        headers.append(label)
    return headers


def source_value(
    record: list[Any],
    headers: list[str],
    header_name: str,
) -> Any:
    try:
        index = next(
            index
            for index, header in enumerate(headers)
            if header.casefold() == header_name.casefold()
        )
    except StopIteration:
        return None
    return record[index] if index < len(record) else None


def first_populated(*values: Any) -> Any:
    return next(
        (value for value in values if value is not None and cell_text(value)),
        None,
    )


def style_sheet(
    worksheet: openpyxl.worksheet.worksheet.Worksheet,
    headers: list[str],
    row_count: int,
    table_name: str,
    title: str,
    subtitle: str,
    sections: list[tuple[int, int, str, str]],
    create_table: bool = True,
) -> None:
    last_column = len(headers)
    last_letter = openpyxl.utils.get_column_letter(last_column)
    worksheet.freeze_panes = "C5" if last_column > 2 else "A5"
    worksheet.sheet_view.showGridLines = False
    worksheet.sheet_properties.tabColor = BLUE
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.print_title_rows = "1:4"
    worksheet.sheet_view.zoomScale = 85

    worksheet.merge_cells(f"A1:{last_letter}1")
    title_cell = worksheet["A1"]
    title_cell.value = title
    title_cell.font = Font(name="Aptos Display", bold=True, color="FFFFFF", size=18)
    title_cell.fill = PatternFill("solid", fgColor=NAVY)
    title_cell.alignment = Alignment(vertical="center", indent=1)
    worksheet.row_dimensions[1].height = 38

    worksheet.merge_cells(f"A2:{last_letter}2")
    subtitle_cell = worksheet["A2"]
    subtitle_cell.value = subtitle
    subtitle_cell.font = Font(name="Aptos", color=MUTED, size=10, italic=True)
    subtitle_cell.fill = PatternFill("solid", fgColor="F2F5F9")
    subtitle_cell.alignment = Alignment(vertical="center", indent=1)
    worksheet.row_dimensions[2].height = 24

    for start_column, end_column, label, fill_color in sections:
        start_letter = openpyxl.utils.get_column_letter(start_column)
        end_letter = openpyxl.utils.get_column_letter(end_column)
        if start_column != end_column:
            worksheet.merge_cells(f"{start_letter}3:{end_letter}3")
        section_cell = worksheet.cell(row=3, column=start_column, value=label)
        section_cell.font = Font(name="Aptos", bold=True, color="FFFFFF", size=9)
        section_cell.fill = PatternFill("solid", fgColor=fill_color)
        section_cell.alignment = Alignment(vertical="center", indent=1)
        for column in range(start_column, end_column + 1):
            worksheet.cell(row=3, column=column).fill = PatternFill(
                "solid", fgColor=fill_color
            )
    worksheet.row_dimensions[3].height = 22

    section_colors = {
        column: fill_color
        for start_column, end_column, _, fill_color in sections
        for column in range(start_column, end_column + 1)
    }
    worksheet.row_dimensions[4].height = 42
    for column_index, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=4, column=column_index, value=header)
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF", size=9)
        cell.fill = PatternFill(
            "solid", fgColor=section_colors.get(column_index, BLUE)
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    narrative_terms = (
        "descrip",
        "motivo",
        "justific",
        "denomin",
        "texto",
        "observ",
        "condici",
        "procedencia",
    )
    narrative_columns: set[int] = set()
    for column_index, header in enumerate(headers, start=1):
        label = header.casefold()
        column_letter = openpyxl.utils.get_column_letter(column_index)
        if any(term in label for term in narrative_terms):
            worksheet.column_dimensions[column_letter].width = 38
            narrative_columns.add(column_index)
        elif "resultado del cruce" in label:
            worksheet.column_dimensions[column_letter].width = 22
        else:
            worksheet.column_dimensions[column_letter].width = max(
                12, min(len(header) + 3, 26)
            )
        if header in {
            "N.º DOCUMENTO CONTABLE SAP",
            "CUENTA CONTABLE SAP",
        }:
            worksheet.column_dimensions[column_letter].width = 25
            for row_index in range(5, row_count + 5):
                worksheet.cell(row=row_index, column=column_index).number_format = "@"
        elif header == "IMPORTE CONTABLE SAP":
            worksheet.column_dimensions[column_letter].width = 22
            for row_index in range(5, row_count + 5):
                worksheet.cell(row=row_index, column=column_index).number_format = (
                    '#,##0.00;[Red]-#,##0.00'
                )
        elif header == "FECHA CONTABILIZACION SAP":
            worksheet.column_dimensions[column_letter].width = 24
            for row_index in range(5, row_count + 5):
                worksheet.cell(row=row_index, column=column_index).number_format = (
                    "dd/mm/yyyy"
                )

    bottom_border = Border(
        bottom=Side(style="hair", color="D9E2F0")
    )
    for row in worksheet.iter_rows(min_row=5, max_row=row_count + 4):
        worksheet.row_dimensions[row[0].row].height = 32
        for cell in row:
            cell.font = Font(name="Aptos", size=10, color=TEXT)
            cell.alignment = Alignment(
                vertical="center",
                horizontal="left" if cell.column in narrative_columns else "center",
                wrap_text=cell.column in narrative_columns,
            )
            cell.border = bottom_border

    if row_count and create_table:
        table = Table(
            displayName=table_name,
            ref=f"A4:{last_letter}{row_count + 4}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)
    elif row_count:
        worksheet.auto_filter.ref = f"A4:{last_letter}{row_count + 4}"


def add_records_sheet(
    workbook: openpyxl.Workbook,
    title: str,
    headers: list[str],
    records: list[list[Any]],
    table_name: str,
    subtitle: str,
    sections: list[tuple[int, int, str, str]],
    create_table: bool = True,
) -> None:
    worksheet = workbook.create_sheet(title=title)
    for _ in range(4):
        worksheet.append([None] * len(headers))
    for record in records:
        worksheet.append(record)
    style_sheet(
        worksheet,
        headers,
        len(records),
        table_name,
        title,
        subtitle,
        sections,
        create_table,
    )


def save_template_cells(
    template_path: Path,
    output_path: Path,
    sheet_name: str,
    updates: list[tuple[str, Any, bool]],
    new_rows: list[tuple[int, dict[int, tuple[Any, bool]]]],
    style_source_row: int,
    filter_header_row: int | None = None,
) -> None:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    relationship_ns = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    package_relationship_ns = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    namespace = {"main": main_ns}

    with ZipFile(template_path, "r") as source_zip:
        workbook_xml = ET.fromstring(source_zip.read("xl/workbook.xml"))
        relationships_xml = ET.fromstring(
            source_zip.read("xl/_rels/workbook.xml.rels")
        )
        sheet = next(
            (
                node
                for node in workbook_xml.findall("main:sheets/main:sheet", namespace)
                if node.get("name") == sheet_name
            ),
            None,
        )
        if sheet is None:
            raise ValueError(f"No existe la hoja {sheet_name} en la plantilla.")
        relationship_id = sheet.get(f"{{{relationship_ns}}}id")
        relationship = next(
            (
                node
                for node in relationships_xml.findall(
                    f"{{{package_relationship_ns}}}Relationship"
                )
                if node.get("Id") == relationship_id
            ),
            None,
        )
        if relationship is None or relationship.get("Target") is None:
            raise ValueError(f"No se encontró la relación XML de {sheet_name}.")
        target = relationship.get("Target", "")
        sheet_path = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )

        sheet_xml = source_zip.read(sheet_path)
        styles_xml = source_zip.read("xl/styles.xml")
        fills_match = re.search(rb"<fills\b[^>]*>(.*?)</fills>", styles_xml, re.DOTALL)
        formats_match = re.search(
            rb"<cellXfs\b([^>]*)>(.*?)</cellXfs>", styles_xml, re.DOTALL
        )
        if fills_match is None or formats_match is None:
            raise ValueError("La plantilla STATUS no contiene estilos de celdas.")
        fills_xml = fills_match.group(1)
        fill_tags = re.findall(rb"<fill\b[^>]*>.*?</fill>", fills_xml, re.DOTALL)
        yellow_fill_id = next(
            (
                index
                for index, fill in enumerate(fill_tags)
                if re.search(rb"<patternFill\b[^>]*\bpatternType=\"solid\"", fill)
                and re.search(
                    rb"<fgColor\b[^>]*\brgb=\"(?:FF)?FFFF00\"", fill, re.IGNORECASE
                )
            ),
            None,
        )
        if yellow_fill_id is None:
            yellow_fill_id = len(fill_tags)
            fill_xml = (
                b'<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/>'
                b'<bgColor indexed="64"/></patternFill></fill>'
            )
            fills_xml += fill_xml
            fills_open = re.search(rb"<fills\b[^>]*>", styles_xml)
            if fills_open is None:
                raise ValueError("No se encontró la sección fills de los estilos.")
            fills_open_xml = re.sub(
                rb"\bcount=\"\d+\"",
                f'count="{len(fill_tags) + 1}"'.encode(),
                fills_open.group(0),
                count=1,
            )
            if fills_open_xml == fills_open.group(0):
                raise ValueError("No se pudo actualizar el conteo de estilos.")
            styles_xml = (
                styles_xml[: fills_match.start()]
                + fills_open_xml
                + fills_xml
                + b"</fills>"
                + styles_xml[fills_match.end() :]
            )
            formats_match = re.search(
                rb"<cellXfs\b([^>]*)>(.*?)</cellXfs>", styles_xml, re.DOTALL
            )
            if formats_match is None:
                raise ValueError("No se pudo actualizar la sección de estilos.")

        cell_formats: list[bytes] = []
        cursor = 0
        while start_match := re.search(rb"<xf\b[^>]*>", formats_match.group(2)[cursor:]):
            start = cursor + start_match.start()
            start_end = cursor + start_match.end()
            if formats_match.group(2)[start:start_end].endswith(b"/>"):
                cell_formats.append(formats_match.group(2)[start:start_end])
                cursor = start_end
                continue
            closing = re.search(rb"</xf\s*>", formats_match.group(2)[start_end:])
            if closing is None:
                raise ValueError("Formato de celda incompleto en styles.xml.")
            end = start_end + closing.end()
            cell_formats.append(formats_match.group(2)[start:end])
            cursor = end
        if not cell_formats:
            raise ValueError("La plantilla no contiene formatos de celda.")
        style_by_original: dict[tuple[int, bool], int] = {}
        appended_formats: list[bytes] = []
        if new_rows:
            style_row_match = re.search(
                rb"<row\b(?=[^>]*\br=\""
                + str(style_source_row).encode()
                + rb"\")[^>]*>.*?</row>",
                sheet_xml,
                re.DOTALL,
            )
            if style_row_match is None:
                raise ValueError(
                    f"No existe la fila de formato {style_source_row} en "
                    f"{sheet_name}."
                )
            style_ids: dict[int, int] = {}
            for cell_match in re.finditer(
                rb"<c\b([^>]*\br=\"([A-Z]+)\d+\"[^>]*)>",
                style_row_match.group(0),
            ):
                style_id_match = re.search(rb"\bs=\"(\d+)\"", cell_match.group(1))
                if style_id_match is not None:
                    style_ids[
                        openpyxl.utils.column_index_from_string(
                            cell_match.group(2).decode("ascii")
                        )
                    ] = int(style_id_match.group(1))

            def make_cell(
                coordinate: str,
                style_id: int,
                value: Any = None,
                is_numeric: bool = False,
            ) -> bytes:
                style_attr = f' s="{style_id}"' if style_id else ""
                if value is None:
                    return f'<c r="{coordinate}"{style_attr}/>'.encode()
                if isinstance(value, datetime):
                    return (
                        f'<c r="{coordinate}"{style_attr}>'
                        f"<v>{excel_serial(value)}</v></c>"
                    ).encode()
                if is_numeric:
                    number = status_amount(value)
                    return (
                        f'<c r="{coordinate}"{style_attr}>'
                        f"<v>{format(number, 'f')}</v></c>"
                    ).encode()
                value_text = escape(str(value))
                preserve = (
                    ' xml:space="preserve"'
                    if str(value) != str(value).strip()
                    else ""
                )
                return (
                    f'<c r="{coordinate}"{style_attr} t="inlineStr">'
                    f"<is><t{preserve}>{value_text}</t></is></c>"
                ).encode("utf-8")

            source_open_tag = re.match(rb"<row\b[^>]*>", style_row_match.group(0))
            if source_open_tag is None:
                raise ValueError("La fila de formato de STATUS no es válida.")
            row_open_tag = source_open_tag.group(0)
            for row_number, values in new_rows:
                row_tag = re.sub(
                    rb'\br="\d+"',
                    f'r="{row_number}"'.encode(),
                    row_open_tag,
                    count=1,
                )
                cell_tags = []
                for column_number in range(1, 37):
                    coordinate = (
                        f"{openpyxl.utils.get_column_letter(column_number)}"
                        f"{row_number}"
                    )
                    value_entry = values.get(column_number)
                    style_id = style_ids.get(column_number, 0)
                    if value_entry is None and style_id == 0:
                        continue
                    value, numeric = value_entry if value_entry is not None else (None, False)
                    cell_tags.append(
                        make_cell(coordinate, style_id, value, numeric)
                    )
                for column_number, formula in (
                    (
                        29,
                        f'+IF(B{row_number}=F{row_number}, "NUEVO","ANTIGUO")',
                    ),
                    (
                        30,
                        f'IF(AB{row_number}="BAJA", "EGRESO",'
                        f'IF(AC{row_number}="NUEVO","INGRESO","SALDO"))',
                    ),
                    (
                        31,
                        f'IF(V{row_number}="SI","TRABAJADO","NO TRABAJADO")',
                    ),
                ):
                    coordinate = (
                        f"{openpyxl.utils.get_column_letter(column_number)}"
                        f"{row_number}"
                    )
                    style_id = style_ids.get(column_number, 0)
                    formula_style_attr = (
                        f' s="{style_id}"' if style_id else ""
                    )
                    formula_cell = (
                        f'<c r="{coordinate}"'
                        f"{formula_style_attr}>"
                        f"<f>{escape(formula)}</f></c>"
                    ).encode("utf-8")
                    cell_tags = [
                        cell for cell in cell_tags
                        if not re.search(
                            rb'\br="' + coordinate.encode() + rb'"',
                            cell,
                        )
                    ]
                    cell_tags.append(formula_cell)
                cell_tags.sort(
                    key=lambda cell: openpyxl.utils.column_index_from_string(
                        re.search(rb'\br="([A-Z]+)', cell).group(1).decode()
                    )
                )
                new_row_xml = (
                    row_tag + b"".join(cell_tags) + b"</row>"
                )
                existing_row = re.search(
                    rb"<row\b(?=[^>]*\br=\""
                    + str(row_number).encode()
                    + rb"\")[^>]*/>"
                    rb"|<row\b(?=[^>]*\br=\""
                    + str(row_number).encode()
                    + rb"\")[^>]*>.*?</row>",
                    sheet_xml,
                    re.DOTALL,
                )
                if existing_row is not None:
                    sheet_xml = (
                        sheet_xml[: existing_row.start()]
                        + new_row_xml
                        + sheet_xml[existing_row.end() :]
                    )
                else:
                    # Insertar antes de la primera fila con número mayor
                    # (Excel exige filas en orden ascendente).
                    insert_at = next(
                        (
                            match.start()
                            for match in re.finditer(
                                rb'<row\b[^>]*\br="(\d+)"', sheet_xml
                            )
                            if int(match.group(1)) > row_number
                        ),
                        sheet_xml.rfind(b"</sheetData>"),
                    )
                    if insert_at < 0:
                        raise ValueError("No se encontró sheetData en STATUS.")
                    sheet_xml = (
                        sheet_xml[:insert_at]
                        + new_row_xml
                        + sheet_xml[insert_at:]
                    )
            last_added_row = max(row_number for row_number, _ in new_rows)
            dimension_match = re.search(rb"<dimension\b[^>]*>", sheet_xml)
            if dimension_match is not None:
                dimension_xml = re.sub(
                    rb":AJ\d+",
                    f":AJ{last_added_row}".encode(),
                    dimension_match.group(0),
                )
                sheet_xml = (
                    sheet_xml[: dimension_match.start()]
                    + dimension_xml
                    + sheet_xml[dimension_match.end() :]
                )
        for coordinate, value, is_numeric in updates:
            match = re.fullmatch(r"([A-Z]+)(\d+)", coordinate)
            if match is None:
                raise ValueError(f"Referencia de celda inválida: {coordinate}")
            column_name, row_text = match.groups()
            row_number = int(row_text)
            row_match = re.search(
                rb"<row\b(?=[^>]*\br=\"" + row_text.encode() + rb"\")[^>]*>.*?</row>",
                sheet_xml,
                re.DOTALL,
            )
            if row_match is None:
                raise ValueError(f"No existe la fila {row_number} en {sheet_name}.")
            row_xml = row_match.group(0)
            cell_pattern = (
                rb"<c\b(?=[^>]*\br=\"" + re.escape(coordinate.encode()) + rb"\")[^>]*/>"
                rb"|<c\b(?=[^>]*\br=\"" + re.escape(coordinate.encode()) + rb"\")[^>]*>.*?</c>"
            )
            existing_cell = re.search(cell_pattern, row_xml, re.DOTALL)
            if existing_cell is None:
                original_style_id = 0
            else:
                style_match = re.search(rb"\bs=\"(\d+)\"", existing_cell.group(0))
                original_style_id = int(style_match.group(1)) if style_match else 0
            if original_style_id >= len(cell_formats):
                raise ValueError(
                    f"Estilo de celda inválido en {coordinate}: {original_style_id}"
                )
            is_date = isinstance(value, datetime)
            style_key = (original_style_id, is_date)
            if style_key not in style_by_original:
                original_format = cell_formats[original_style_id]
                if is_date:
                    # 17 = formato integrado "mmm-yy" (igual que MES DE BAJA).
                    original_format = re.sub(
                        rb"\s(?:numFmtId|applyNumberFormat)=\"[^\"]*\"",
                        b"",
                        original_format,
                    )
                    original_format = re.sub(
                        rb"<xf\b",
                        b'<xf numFmtId="17" applyNumberFormat="1"',
                        original_format,
                        count=1,
                    )
                fill_match = re.search(rb"\bfillId=\"(\d+)\"", original_format)
                current_fill_id = int(fill_match.group(1)) if fill_match else 0
                apply_fill = re.search(rb"\bapplyFill=\"([^\"]+)\"", original_format)
                if (
                    current_fill_id == yellow_fill_id
                    and apply_fill is not None
                    and apply_fill.group(1) == b"1"
                    and not is_date
                ):
                    style_by_original[style_key] = original_style_id
                else:
                    updated_format, replacements = re.subn(
                        rb"\bfillId=\"\d+\"",
                        f'fillId="{yellow_fill_id}"'.encode(),
                        original_format,
                        count=1,
                    )
                    if replacements == 0:
                        updated_format = re.sub(
                            rb"<xf\b",
                            f'<xf fillId="{yellow_fill_id}"'.encode(),
                            original_format,
                            count=1,
                        )
                    if apply_fill is not None:
                        updated_format = re.sub(
                            rb"\bapplyFill=\"[^\"]+\"",
                            b'applyFill="1"',
                            updated_format,
                            count=1,
                        )
                    else:
                        updated_format = re.sub(
                            rb"<xf\b",
                            b'<xf applyFill="1"',
                            updated_format,
                            count=1,
                        )
                    style_by_original[style_key] = (
                        len(cell_formats) + len(appended_formats)
                    )
                    appended_formats.append(updated_format)
            style_id = style_by_original[style_key]
            if is_date:
                cell_xml = (
                    f'<c r="{coordinate}" s="{style_id}">'
                    f"<v>{excel_serial(value)}</v></c>"
                ).encode("utf-8")
            elif is_numeric:
                number = status_amount(value)
                cell_xml = (
                    f'<c r="{coordinate}" s="{style_id}">'
                    f"<v>{format(number, 'f')}</v></c>"
                ).encode("utf-8")
            else:
                value_text = escape(str(value))
                preserve = (
                    ' xml:space="preserve"'
                    if str(value) != str(value).strip()
                    else ""
                )
                cell_xml = (
                    f'<c r="{coordinate}" s="{style_id}" t="inlineStr">'
                    f"<is><t{preserve}>{value_text}</t></is></c>"
                ).encode("utf-8")
            if existing_cell is not None:
                row_xml = (
                    row_xml[: existing_cell.start()]
                    + cell_xml
                    + row_xml[existing_cell.end() :]
                )
            else:
                column_number = openpyxl.utils.column_index_from_string(column_name)
                cell_tags = list(
                    re.finditer(rb"<c\b[^>]*\br=\"([A-Z]+)\d+\"", row_xml)
                )
                insert_at = next(
                    (
                        cell.start()
                        for cell in cell_tags
                        if openpyxl.utils.column_index_from_string(
                            cell.group(1).decode("ascii")
                        )
                        > column_number
                    ),
                    row_xml.rfind(b"</row>"),
                )
                row_xml = row_xml[:insert_at] + cell_xml + row_xml[insert_at:]
            sheet_xml = (
                sheet_xml[: row_match.start()]
                + row_xml
                + sheet_xml[row_match.end() :]
            )

        if appended_formats:
            formats_xml_content = formats_match.group(2) + b"".join(appended_formats)
            attributes = re.sub(
                rb"\bcount=\"\d+\"",
                f'count="{len(cell_formats) + len(appended_formats)}"'.encode(),
                formats_match.group(1),
                count=1,
            )
            formats_replacement = (
                b"<cellXfs"
                + (b" " + attributes if attributes else b"")
                + b">"
                + formats_xml_content
                + b"</cellXfs>"
            )
            styles_xml = (
                styles_xml[: formats_match.start()]
                + formats_replacement
                + styles_xml[formats_match.end() :]
            )

        # Filtro único sobre toda la hoja: desde los encabezados hasta la
        # última fila con datos, como si fuera una sola tabla.
        workbook_xml_bytes = source_zip.read("xl/workbook.xml")
        if filter_header_row is not None:
            last_filter_row = max(
                [style_source_row] + [row_number for row_number, _ in new_rows]
            )
            last_column = openpyxl.utils.get_column_letter(len(STATUS_HEADERS))
            filter_ref = f"A{filter_header_row}:{last_column}{last_filter_row}"
            if re.search(rb"<autoFilter\b", sheet_xml):
                sheet_xml = re.sub(
                    rb'(<autoFilter\b[^>]*\bref=")[^"]+(")',
                    rb"\g<1>" + filter_ref.encode() + rb"\g<2>",
                    sheet_xml,
                    count=1,
                )
            else:
                insert_match = re.search(
                    rb"</sheetData>(?:\s*<sheetCalcPr\b[^>]*/>)?"
                    rb"(?:\s*<sheetProtection\b[^>]*/>)?",
                    sheet_xml,
                )
                if insert_match is None:
                    raise ValueError("No se encontró sheetData en STATUS.")
                sheet_xml = (
                    sheet_xml[: insert_match.end()]
                    + f'<autoFilter ref="{filter_ref}"/>'.encode()
                    + sheet_xml[insert_match.end() :]
                )
            sheet_index = next(
                index
                for index, node in enumerate(
                    workbook_xml.findall("main:sheets/main:sheet", namespace)
                )
                if node.get("name") == sheet_name
            )
            absolute_ref = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", filter_ref)
            quoted_name = "'" + sheet_name.replace("'", "''") + "'"
            filter_name = (
                f'<definedName name="_xlnm._FilterDatabase" '
                f'localSheetId="{sheet_index}" hidden="1">'
                f"{escape(quoted_name)}!{absolute_ref}</definedName>"
            ).encode("utf-8")
            existing_name = re.search(
                rb'<definedName\b(?=[^>]*name="_xlnm\._FilterDatabase")'
                rb'(?=[^>]*localSheetId="' + str(sheet_index).encode() + rb'")'
                rb"[^>]*>[^<]*</definedName>",
                workbook_xml_bytes,
            )
            if existing_name is not None:
                workbook_xml_bytes = (
                    workbook_xml_bytes[: existing_name.start()]
                    + filter_name
                    + workbook_xml_bytes[existing_name.end() :]
                )
            elif b"<definedNames>" in workbook_xml_bytes:
                workbook_xml_bytes = workbook_xml_bytes.replace(
                    b"<definedNames>", b"<definedNames>" + filter_name, 1
                )
            else:
                workbook_xml_bytes = re.sub(
                    rb"(</sheets>)",
                    rb"\1<definedNames>" + filter_name + rb"</definedNames>",
                    workbook_xml_bytes,
                    count=1,
                )

        # Control estructural: Excel marca el archivo como dañado si las
        # filas o celdas no están en orden ascendente o si se repiten.
        previous_row = 0
        for row_match in re.finditer(
            rb'<row\b[^>]*\br="(\d+)"[^>]*?(?:/>|>(.*?)</row>)',
            sheet_xml,
            re.DOTALL,
        ):
            current_row = int(row_match.group(1))
            if current_row <= previous_row:
                raise ValueError(
                    f"Filas fuera de orden en {sheet_name}: {current_row} "
                    f"después de {previous_row}."
                )
            previous_row = current_row
            columns = [
                openpyxl.utils.column_index_from_string(column.decode())
                for column in re.findall(
                    rb'<c\b[^>]*\br="([A-Z]+)\d+"', row_match.group(2) or b""
                )
            ]
            if columns != sorted(set(columns)):
                raise ValueError(
                    f"Celdas fuera de orden o repetidas en la fila {current_row}."
                )
        if b"<v></v>" in sheet_xml:
            raise ValueError("Hay celdas con valor vacío <v></v> en STATUS.")

        # Que Excel recalcule las fórmulas nuevas al abrir el archivo.
        if re.search(rb"<calcPr\b", workbook_xml_bytes):
            workbook_xml_bytes = re.sub(
                rb"<calcPr\b(?![^>]*fullCalcOnLoad)",
                b'<calcPr fullCalcOnLoad="1"',
                workbook_xml_bytes,
                count=1,
            )
        patched_parts = {
            "xl/styles.xml": styles_xml,
            sheet_path: sheet_xml,
            "xl/workbook.xml": workbook_xml_bytes,
        }
        staging_path = output_path.with_name(
            f".{output_path.stem}.package.tmp.xlsx"
        )
        try:
            with ZipFile(staging_path, "w") as result_zip:
                for entry in source_zip.infolist():
                    result_zip.writestr(
                        entry,
                        patched_parts.get(entry.filename, source_zip.read(entry)),
                    )
            staging_path.replace(output_path)
        finally:
            if staging_path.exists():
                staging_path.unlink()


def update_status_template(
    candidates: list[dict[str, Any]],
    template_path: Path,
    output_folder: Path,
    timestamp: str,
    resolve_equipment: Any,
    ih08_fields: Any,
    asset_book_value: Any,
    asset_denomination: Any,
    baja_documents: dict[str, set[tuple[str, datetime]]],
    hana_from_r3: Any = None,
    legacy_book_info: Any = None,
    r3_legacy_detail: Any = None,
) -> dict[str, Any]:
    """Completa el STATUS sin tocar lo existente.

    - Filas existentes: solo se escriben celdas vacías (CODIGO AF HANAH,
      VALOR CONTABLE, datos de baja desde ACDOCA y datos de IH08).
    - Filas nuevas: registros de la BASE que aún no están en el STATUS
      (por SERIE SAP o por N° de ficha FIC + descripción/placa).
    - Todo lo escrito queda en amarillo. Las celdas con '-' u otro texto
      se consideran llenas y no se tocan.
    """
    if not template_path.is_file():
        raise FileNotFoundError(f"No existe la plantilla STATUS: {template_path}")
    if not output_folder.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de salida: {output_folder}")

    current_month = SPANISH_MONTHS[datetime.now().month - 1]
    workbook = openpyxl.load_workbook(template_path)
    try:
        if STATUS_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f"La plantilla STATUS no contiene la hoja requerida: "
                f"{STATUS_SHEET_NAME}"
            )
        worksheet = workbook[STATUS_SHEET_NAME]
        header_row = next(
            (
                row_number
                for row_number in range(1, min(worksheet.max_row, 10) + 1)
                if any(
                    normalize_header(cell.value) == normalize_header("SERIE SAP")
                    for cell in worksheet[row_number]
                )
            ),
            None,
        )
        if header_row is None:
            raise ValueError(
                f"No se encontró la fila de encabezados en {STATUS_SHEET_NAME}."
            )
        target_columns: dict[str, int] = {}
        for cell in worksheet[header_row]:
            normalized = normalize_header(cell.value)
            if normalized:
                target_columns.setdefault(normalized, cell.column)
        missing_headers = [
            header
            for header in STATUS_HEADERS
            if normalize_header(header) not in target_columns
        ]
        if missing_headers:
            raise ValueError(
                f"La hoja {STATUS_SHEET_NAME} no tiene las columnas: "
                f"{', '.join(missing_headers)}. Revisa STATUS_HEADERS."
            )

        def column(header: str) -> int:
            return target_columns[normalize_header(header)]

        def letter(header: str) -> str:
            return openpyxl.utils.get_column_letter(column(header))

        populated_rows = [
            row_number
            for row_number in range(header_row + 1, worksheet.max_row + 1)
            if any(
                cell.value is not None and cell_text(cell.value)
                for cell in worksheet[row_number]
            )
        ]
        last_data_row = max(populated_rows, default=header_row)

        def existing_value(row_number: int, header: str) -> Any:
            return worksheet.cell(row_number, column(header)).value

        # Índices de lo que ya existe en el STATUS.
        status_equipment: set[str] = set()
        status_fic: dict[str, list[tuple[str, str]]] = {}
        hana_rows: Counter[str] = Counter()
        for row_number in range(header_row + 1, last_data_row + 1):
            equipment = normalize_key(existing_value(row_number, "SERIE SAP"))
            if equipment:
                status_equipment.add(equipment)
            fic = fic_number(existing_value(row_number, "DOC. REFERENCIA"))
            if fic and "fic" in cell_text(
                existing_value(row_number, "DOC. REFERENCIA")
            ).casefold():
                status_fic.setdefault(fic, []).append(
                    (
                        normalize_search_text(
                            existing_value(row_number, "DESCRIPCION DEL ACTIVO")
                        ),
                        normalize_status_plate(existing_value(row_number, "PLACA")),
                    )
                )
            hana = cell_text(existing_value(row_number, "CODIGO AF HANAH"))
            if re.fullmatch(r"\d{12}", hana):
                hana_rows[hana] += 1

        validation_cache: dict[str, dict[str, Any]] = {}

        def resolve(equipment: str) -> dict[str, Any]:
            if not equipment:
                return {}
            if equipment not in validation_cache:
                validation_cache[equipment] = resolve_equipment(equipment)
            return validation_cache[equipment]

        def no_hana_detail(validation: dict[str, Any]) -> str:
            parts = [validation.get("detail", "")]
            legacy = validation.get("legacy") or {}
            entry = legacy.get("entry")
            if entry is not None:
                parts.append(
                    f"Solo en AR02 ANTIGUO: R3 {legacy['key'][0]}-"
                    f"{legacy['key'][1]} '{entry['denom']}', CAP actuales "
                    f"{entry['cap']}, Val.cont.act. {entry['valcont']}"
                    + (f" ({legacy_asset_state(entry)})" if legacy_asset_state(entry) else "")
                    + "."
                )
            elif legacy.get("code"):
                parts.append(f"AF R3 {legacy['code']} no está en el AR02 ANTIGUO.")
            return " ".join(part for part in parts if part)

        def resolve_r3(r3_code: str) -> list[str]:
            if not hana_from_r3 or not re.fullmatch(r"\d{10}", r3_code):
                return []
            return hana_from_r3(r3_code)

        # Pre-paso A: códigos HANAH previstos en filas existentes.
        for row_number in range(header_row + 1, last_data_row + 1):
            if cell_text(existing_value(row_number, "CODIGO AF HANAH")):
                continue
            validation = resolve(
                normalize_key(existing_value(row_number, "SERIE SAP"))
            )
            if validation.get("status") == "VALIDADO AR02":
                hana_rows[validation["anln1"]] += 1
            else:
                r3_hana = resolve_r3(
                    cell_text(existing_value(row_number, "CODIGO AF R3"))
                )
                if len(r3_hana) == 1:
                    hana_rows[r3_hana[0]] += 1

        # Pre-paso B: registros de la BASE que aún no están en el STATUS.
        to_append: list[dict[str, Any]] = []
        skipped_existing = 0
        added_keys: set[tuple[str, str]] = set()
        for candidate in candidates:
            equipment = candidate["equipment"]
            fic = candidate["fic"]
            description = normalize_search_text(candidate["description"])
            plate = normalize_status_plate(candidate["plate"])
            if equipment and equipment in status_equipment:
                skipped_existing += 1
                continue
            fic_rows = status_fic.get(fic, [])
            if fic_rows and any(
                (description and fuzz.token_set_ratio(description, row_desc) >= 90)
                or (plate and is_plate_like(plate) and plate == row_plate)
                for row_desc, row_plate in fic_rows
            ):
                skipped_existing += 1
                continue
            key = (fic, equipment or description)
            if key in added_keys:
                continue
            added_keys.add(key)
            if equipment:
                status_equipment.add(equipment)
            to_append.append(candidate)
            if candidate["validation"].get("status") == "VALIDADO AR02":
                hana_rows[candidate["validation"]["anln1"]] += 1

        updates: list[tuple[str, Any, bool]] = []
        fill_counts: Counter[str] = Counter()
        examples: list[tuple[int, str, Any]] = []
        trace: list[list[Any]] = []
        review: list[list[Any]] = []

        def record_fill(
            row_number: int,
            row_type: str,
            equipment: str,
            header: str,
            value: Any,
            source: str,
        ) -> None:
            numeric = header in {"VALOR CONTABLE", "N° DE DOCUMENTO"}
            updates.append((f"{letter(header)}{row_number}", value, numeric))
            fill_counts[header] += 1
            shown = value.strftime("%m/%Y") if isinstance(value, datetime) else value
            trace.append([row_number, row_type, equipment, header, shown, source])
            if len(examples) < 8:
                examples.append((row_number, header, shown))

        def sap_values(
            hana: str,
            already_filled: Any,
            row_number: int,
            equipment: str,
            existing_document: str = "",
            validation: dict[str, Any] | None = None,
        ) -> list[tuple[str, Any, str]]:
            """Valores derivados del activo (AR02 + ACDOCA) para celdas vacías."""
            values: list[tuple[str, Any, str]] = []
            legacy_info = (
                legacy_book_info(hana, validation) if legacy_book_info else None
            )
            if legacy_info and legacy_info["alert"]:
                review.append(
                    [
                        row_number,
                        equipment,
                        hana,
                        "AR02 HANA EN 0; AR02 ANTIGUO CON VALOR",
                        legacy_info["detail"],
                    ]
                )
            if not already_filled("VALOR CONTABLE"):
                if hana_rows[hana] > 1:
                    review.append(
                        [
                            row_number,
                            equipment,
                            hana,
                            "VALOR CONTABLE NO LLENADO",
                            "El mismo CODIGO AF HANAH aparece en varias filas; "
                            "el valor se reparte manualmente.",
                        ]
                    )
                else:
                    book_value, source = asset_book_value(hana, validation)
                    if book_value is not None:
                        values.append(("VALOR CONTABLE", book_value, source))
            documents = sorted(baja_documents.get(hana, set()))
            if len(documents) > 1:
                review.append(
                    [
                        row_number,
                        equipment,
                        hana,
                        "VARIOS DOCUMENTOS DE BAJA",
                        " / ".join(
                            f"{doc} ({date:%m/%Y})" for doc, date in documents
                        ),
                    ]
                )
            elif documents:
                document, posting_date = documents[0]
                if existing_document and existing_document != document:
                    review.append(
                        [
                            row_number,
                            equipment,
                            hana,
                            "N° DE DOCUMENTO DISTINTO A ACDOCA",
                            f"STATUS: {existing_document} / ACDOCA: {document} "
                            f"({posting_date:%d/%m/%Y}). ¿Error de tipeo?",
                        ]
                    )
                month_name = SPANISH_MONTHS[posting_date.month - 1].lower()
                source = f"ACDOCA baja doc {document} ({posting_date:%d/%m/%Y})"
                values.extend(
                    [
                        ("SE DIO DE BAJA", "Si", source),
                        (
                            "MES DE BAJA",
                            datetime(posting_date.year, posting_date.month, 1),
                            source,
                        ),
                        ("N° DE DOCUMENTO", int(document), source),
                        (
                            "COMENTARIOS",
                            f"Se dió de baja en {month_name}-{posting_date.year}",
                            source,
                        ),
                        ("AUXILIAR  1", "BAJA", source),
                    ]
                )
            return [
                (header, value, source)
                for header, value, source in values
                if not already_filled(header)
            ]

        # 1) Filas existentes: completar solo celdas vacías.
        filled_existing_rows = 0
        for row_number in range(header_row + 1, last_data_row + 1):
            equipment = normalize_key(existing_value(row_number, "SERIE SAP"))
            hana = cell_text(existing_value(row_number, "CODIGO AF HANAH"))
            if (
                not equipment
                and not hana
                and not cell_text(existing_value(row_number, "CODIGO AF R3"))
            ):
                continue

            def already_filled(header: str, row_number: int = row_number) -> bool:
                return bool(cell_text(existing_value(row_number, header)))

            row_fills: list[tuple[str, Any, str]] = []
            validation = resolve(equipment)
            r3_code = cell_text(existing_value(row_number, "CODIGO AF R3"))
            if re.fullmatch(r"\d{12}", hana):
                if (
                    validation.get("status") == "VALIDADO AR02"
                    and validation["anln1"] != hana
                ):
                    review.append(
                        [
                            row_number,
                            equipment,
                            hana,
                            "HANAH DEL STATUS ≠ ACTIVO DE IH08",
                            f"STATUS {hana} ({asset_denomination(hana)}) / IH08 "
                            f"{validation['anln1']} "
                            f"({asset_denomination(validation['anln1'])}). "
                            "No se modificó.",
                        ]
                    )
            elif not already_filled("CODIGO AF HANAH"):
                if validation.get("status") == "VALIDADO AR02":
                    hana = validation["anln1"]
                    row_fills.append(
                        (
                            "CODIGO AF HANAH",
                            hana,
                            f"{validation.get('route', 'IH08')}: equipo {equipment}"
                            f" → {hana} ({asset_denomination(hana)})",
                        )
                    )
                else:
                    r3_hana = resolve_r3(r3_code)
                    if len(r3_hana) == 1:
                        hana = r3_hana[0]
                        row_fills.append(
                            (
                                "CODIGO AF HANAH",
                                hana,
                                f"AR02 AF antiguo: R3 {r3_code} → {hana} "
                                f"({asset_denomination(hana)})",
                            )
                        )
                    else:
                        if validation.get("status"):
                            review.append(
                                [
                                    row_number,
                                    equipment,
                                    "",
                                    f"SIN HANAH: {validation['status']}",
                                    no_hana_detail(validation),
                                ]
                            )
                        if len(r3_hana) > 1:
                            review.append(
                                [
                                    row_number,
                                    equipment,
                                    "",
                                    "CODIGO AF R3 CON VARIOS HANAH",
                                    f"R3 {r3_code} → {', '.join(r3_hana)}.",
                                ]
                            )
                        elif re.fullmatch(r"\d{10}", r3_code):
                            review.append(
                                [
                                    row_number,
                                    equipment,
                                    "",
                                    "TIENE CODIGO AF R3 SIN HANAH",
                                    f"R3 {r3_code}: no figura como 'Activo "
                                    "fijo orig.' en el AR02 HANA (no migrado)."
                                    + (
                                        " " + r3_legacy_detail(r3_code)
                                        if r3_legacy_detail
                                        else ""
                                    ),
                                ]
                            )

            legacy = validation.get("legacy") or {}
            if (
                legacy.get("code")
                and not r3_code
                and not already_filled("CODIGO AF R3")
                and (not re.fullmatch(r"\d{12}", hana) or hana == validation.get("anln1"))
            ):
                row_fills.append(
                    (
                        "CODIGO AF R3",
                        legacy["code"],
                        f"{legacy['source']}: equipo {equipment}",
                    )
                )

            if re.fullmatch(r"\d{12}", hana):
                existing_document = cell_text(
                    existing_value(row_number, "N° DE DOCUMENTO")
                )
                row_fills.extend(
                    sap_values(
                        hana,
                        already_filled,
                        row_number,
                        equipment,
                        existing_document
                        if existing_document.isdigit()
                        else "",
                        validation,
                    )
                )

            ih08 = ih08_fields(equipment) if equipment else None
            if ih08:
                plate = ih08.get("plate")
                for header, value, source in (
                    (
                        "PLACA",
                        normalize_status_plate(plate) if is_plate_like(plate) else "",
                        "IH08 Número de inventario",
                    ),
                    ("MARCA", ih08.get("manufacturer"), "IH08 Fabricante"),
                    ("Nº SERIE", ih08.get("material"), "IH08 Material"),
                    ("CENTRO", ih08.get("center"), "IH08 Centro planificación"),
                ):
                    if cell_text(value) and not already_filled(header):
                        row_fills.append((header, cell_text(value), source))

            for header, value, source in row_fills:
                record_fill(row_number, "EXISTENTE", equipment, header, value, source)
            if row_fills:
                filled_existing_rows += 1

        # 2) Filas nuevas desde la BASE.
        new_rows: list[tuple[int, dict[int, tuple[Any, bool]]]] = []
        appended_base_rows = 0
        append_row = last_data_row + 1
        for candidate in to_append:
            equipment = candidate["equipment"]
            row_number = append_row
            values: dict[str, tuple[Any, str]] = {
                header: (value, f"BASE fila {candidate['base_row']}: {source}")
                for header, (value, source) in candidate["values"].items()
                if value is not None and cell_text(value)
            }
            values["MES TRABAJADO"] = (current_month, "Mes actual (nuevo ingreso)")
            values["MES RECEPCION FICHA DE ESTADO"] = (
                current_month.capitalize(),
                "Mes actual (nuevo ingreso)",
            )
            validation = candidate["validation"]
            hana = ""
            if validation.get("status") == "VALIDADO AR02":
                hana = validation["anln1"]
                values["CODIGO AF HANAH"] = (
                    hana,
                    f"{validation.get('route', 'IH08')}: equipo {equipment} → "
                    f"{hana} ({asset_denomination(hana)})",
                )
            elif equipment:
                review.append(
                    [
                        row_number,
                        equipment,
                        "",
                        f"SIN HANAH: {validation.get('status', '')}",
                        no_hana_detail(validation),
                    ]
                )
            else:
                review.append(
                    [
                        row_number,
                        "",
                        "",
                        "SIN HANAH: FICHA SIN CÓDIGO DE EQUIPO",
                        "Ver hoja CANDIDATOS AR02 (sugerencias por descripción).",
                    ]
                )
            legacy = validation.get("legacy") or {}
            if legacy.get("code"):
                values["CODIGO AF R3"] = (
                    legacy["code"],
                    f"{legacy['source']}: equipo {equipment}",
                )
            if hana:
                for header, value, source in sap_values(
                    hana,
                    lambda header: header in values,
                    row_number,
                    equipment,
                    "",
                    validation,
                ):
                    values[header] = (value, source)

            row_values: dict[int, tuple[Any, bool]] = {}
            for header in STATUS_HEADERS:
                if header not in values:
                    continue
                value, source = values[header]
                row_values[column(header)] = (
                    value,
                    header in {"VALOR CONTABLE", "N° DE DOCUMENTO"},
                )
                record_fill(
                    row_number,
                    f"NUEVA (BASE fila {candidate['base_row']})",
                    equipment,
                    header,
                    value,
                    source,
                )
            new_rows.append((row_number, row_values))
            appended_base_rows += 1
            append_row += 1

        output_path = output_folder / f"Status_baja_actualizado_{timestamp}.xlsx"
        suffix = 1
        while output_path.exists():
            output_path = output_folder / (
                f"Status_baja_actualizado_{timestamp}_{suffix:02d}.xlsx"
            )
            suffix += 1
    finally:
        workbook.close()

    save_template_cells(
        template_path,
        output_path,
        STATUS_SHEET_NAME,
        updates,
        new_rows,
        last_data_row,
        header_row,
    )
    changed_cells = verify_status_unchanged(
        template_path, output_path, STATUS_SHEET_NAME, last_data_row
    )
    if changed_cells:
        raise ValueError(
            "Se modificaron celdas que ya tenían información en el STATUS: "
            f"{changed_cells[:10]}"
        )

    return {
        "output_path": output_path,
        "fill_counts": fill_counts,
        "examples": examples,
        "filled_cells": sum(fill_counts.values()),
        "filled_existing_rows": filled_existing_rows,
        "appended_rows": appended_base_rows,
        "skipped_existing": skipped_existing,
        "trace": trace,
        "review": review,
        "verified_cells": last_data_row,
    }


def verify_status_unchanged(
    template_path: Path,
    output_path: Path,
    sheet_name: str,
    last_data_row: int,
) -> list[str]:
    """Compara celda por celda: ninguna celda con dato previo puede cambiar."""
    original = openpyxl.load_workbook(template_path)[sheet_name]
    result = openpyxl.load_workbook(output_path)[sheet_name]
    changed: list[str] = []
    for row in original.iter_rows(min_row=1, max_row=last_data_row):
        for cell in row:
            if cell.value is None or cell_text(cell.value) == "":
                continue
            new_value = result[cell.coordinate].value
            if new_value != cell.value:
                changed.append(cell.coordinate)
    return changed


def generate() -> Path:
    base_path = newest_excel(BASE_FOLDER)
    export_path = newest_excel(EXPORT_FOLDER)
    ar02_path = newest_excel(AR02_FOLDER)
    acdoca_path = newest_excel(ACDOCA_FOLDER)
    bkpf_paths = newest_excel_parts(BKPF_FOLDER)

    ar02_headers, ar02_records, ar02_column_indices = read_ar02(ar02_path)
    ar02_header_indices = {
        normalize_header(header): index
        for index, header in enumerate(ar02_headers)
    }
    ar02_value_indices: list[int | None] = []
    for header in AR02_VALUE_HEADERS:
        index = ar02_header_indices.get(normalize_header(header))
        if index is None and header in AR02_OPTIONAL_VALUE_HEADERS:
            ar02_value_indices.append(None)
            continue
        if index is None:
            raise ValueError(
                f"Falta la columna {header!r} requerida para los valores "
                f"del AR02 en {ar02_path.name}."
            )
        ar02_value_indices.append(index)
    ar02_by_key: dict[tuple[str, str], list[tuple[int, list[Any]]]] = {}
    ar02_by_asset: dict[str, list[tuple[int, list[Any]]]] = {}
    for source_row, record in ar02_records:
        asset_code = normalize_asset_code(record[ar02_column_indices["anln1"]])
        subnumber = normalize_asset_subnumber(
            record[ar02_column_indices["anln2"]]
        )
        ar02_by_key.setdefault((asset_code, subnumber), []).append(
            (source_row, record)
        )
        ar02_by_asset.setdefault(asset_code, []).append((source_row, record))
    # Equivalencias R3 → HANA según la columna de activo antiguo del AR02.
    ar02_by_legacy: dict[str, set[tuple[str, str]]] = {}
    legacy_by_hana: dict[tuple[str, str], tuple[str, str]] = {}
    legacy_column = ar02_column_indices.get("legacy")
    if legacy_column is not None:
        for source_row, record in ar02_records:
            legacy_code, legacy_sub = normalize_legacy_asset(record[legacy_column])
            if not legacy_code:
                continue
            hana_key = (
                normalize_asset_code(record[ar02_column_indices["anln1"]]),
                normalize_asset_subnumber(record[ar02_column_indices["anln2"]]),
            )
            ar02_by_legacy.setdefault(legacy_code, set()).add(hana_key)
            legacy_by_hana[hana_key] = (legacy_code, legacy_sub or hana_key[1])
    ar02_plate_indices = [
        ar02_header_indices[normalize_header(header)]
        for header in AR02_PLATE_HEADERS
        if normalize_header(header) in ar02_header_indices
    ]
    ar02_by_plate: dict[str, set[tuple[str, str]]] = {}
    for source_row, record in ar02_records:
        hana_key = (
            normalize_asset_code(record[ar02_column_indices["anln1"]]),
            normalize_asset_subnumber(record[ar02_column_indices["anln2"]]),
        )
        for index in ar02_plate_indices:
            for plate in split_plates(record[index]):
                ar02_by_plate.setdefault(plate, set()).add(hana_key)
    ar02_legacy_path, ar02_legacy = read_ar02_legacy(AR02_LEGACY_FOLDER)
    ar02_legacy_by_code: dict[str, list[tuple[str, str]]] = {}
    for legacy_key in ar02_legacy:
        ar02_legacy_by_code.setdefault(legacy_key[0], []).append(legacy_key)

    # Índice nombre + centro (solo subnúmero 0000) para la ruta alternativa
    # cuando el activo que trae IH08 está en otro centro.
    ar02_by_name_center: dict[tuple[str, str], list[tuple[str, str]]] = {}
    for (asset_code, subnumber), rows in ar02_by_key.items():
        if subnumber != "0000" or len(rows) != 1:
            continue
        record = rows[0][1]
        ar02_by_name_center.setdefault(
            (
                normalize_search_text(record[ar02_column_indices["denom"]]),
                normalize_key(record[ar02_column_indices["center"]]),
            ),
            [],
        ).append((asset_code, subnumber))

    base_workbook =openpyxl.load_workbook(base_path, read_only=True, data_only=True)
    base_sheet = next(
        (
            sheet
            for sheet in base_workbook.worksheets
            if BASE_SHEET_MARKER in sheet.title.upper()
        ),
        None,
    )
    if base_sheet is None:
        base_workbook.close()
        raise ValueError(
            f"No se encontró una hoja con '{BASE_SHEET_MARKER}' en {base_path.name}."
        )

    base_rows = list(
        base_sheet.iter_rows(min_row=BASE_HEADER_ROW, values_only=True)
    )
    base_headers = unique_headers(base_rows[0])
    base_records: list[list[Any]] = []
    base_source_rows: list[int] = []
    for source_row, row in enumerate(
        base_rows[1:], start=BASE_HEADER_ROW + 1
    ):
        if any(value is not None and str(value).strip() for value in row):
            base_records.append(list(row))
            base_source_rows.append(source_row)
    base_workbook.close()

    export_workbook = openpyxl.load_workbook(
        export_path, read_only=True, data_only=True
    )
    export_sheet = export_workbook.active
    export_rows = list(export_sheet.iter_rows(values_only=True))
    if not export_rows:
        export_workbook.close()
        raise ValueError(f"El archivo de exportación está vacío: {export_path.name}.")

    export_headers = unique_headers(export_rows[0], prefix="Campo IH08")
    normalized_export_headers = [
        normalize_header(header) for header in export_headers
    ]
    validity_matches = [
        index
        for index, header in enumerate(normalized_export_headers)
        if header == "fin de validez"
    ]
    if len(validity_matches) != 1:
        export_workbook.close()
        raise ValueError(
            f"Se esperaba una columna única 'Fin de validez' en "
            f"{export_path.name}; se encontraron {len(validity_matches)}."
        )
    validity_index = validity_matches[0]
    ih08_plate_matches = find_legacy_header(
        normalized_export_headers, IH08_LEGACY_PLATE_HEADERS
    )
    if len(ih08_plate_matches) > 1:
        export_workbook.close()
        raise ValueError(
            "Hay más de una columna de placa antigua en "
            f"{export_path.name}: "
            f"{', '.join(export_headers[index] for index in ih08_plate_matches)}."
        )
    ih08_plate_index = ih08_plate_matches[0] if ih08_plate_matches else None
    export_records_with_rows = [
        (source_row, list(row))
        for source_row, row in enumerate(export_rows[1:], start=2)
        if any(value is not None and str(value).strip() for value in row)
        and validity_index < len(row)
        and is_current_ih08_row(row[validity_index])
    ]
    ih08_equipment_seen = {
        normalize_key(row[EXPORT_EQUIPMENT_COLUMN - 1])
        for row in export_rows[1:]
        if len(row) >= EXPORT_EQUIPMENT_COLUMN
        and normalize_key(row[EXPORT_EQUIPMENT_COLUMN - 1])
    }
    export_records = [record for _, record in export_records_with_rows]
    export_workbook.close()
    if not export_records:
        raise ValueError(
            f"No hay filas IH08 vigentes con Fin de validez = 31.12.9999 "
            f"en {export_path.name}."
        )

    acdoca_headers, acdoca_records, acdoca_column_indices = read_acdoca(
        acdoca_path
    )
    _, bkpf_records, bkpf_column_indices = read_bkpf(bkpf_paths)
    acdoca_by_asset_key: dict[
        tuple[str, str], list[tuple[int, list[Any]]]
    ] = {}
    acdoca_by_asset_2026: dict[
        tuple[str, str], list[tuple[int, list[Any]]]
    ] = {}
    acdoca_by_document: dict[
        tuple[str, str, str], list[tuple[int, list[Any]]]
    ] = {}
    acdoca_without_asset = 0
    for source_row, record in acdoca_records:
        asset_code = normalize_asset_code(
            record[acdoca_column_indices["anln1"]]
        )
        subnumber = normalize_asset_subnumber(
            record[acdoca_column_indices["anln2"]]
        ) or "0000"
        document_class_index = acdoca_column_indices.get("blart")
        document_class = (
            normalize_key(record[document_class_index]).upper()
            if document_class_index is not None
            else ""
        )
        if not asset_code:
            acdoca_without_asset += 1
        elif document_class != "AF":
            asset_key = (asset_code, subnumber)
            acdoca_by_asset_key.setdefault(asset_key, []).append(
                (source_row, record)
            )
            if normalize_key(record[acdoca_column_indices["gjahr"]]) == "2026":
                acdoca_by_asset_2026.setdefault(asset_key, []).append(
                    (source_row, record)
                )
        document_key = (
            normalize_key(record[acdoca_column_indices["rbukrs"]]),
            normalize_key(record[acdoca_column_indices["gjahr"]]),
            normalize_numeric_identifier(
                record[acdoca_column_indices["belnr"]], 10
            ),
        )
        acdoca_by_document.setdefault(document_key, []).append(
            (source_row, record)
        )
    baseline_acdoca_asset_codes = {
        normalize_asset_code(record[acdoca_column_indices["anln1"]])
        for _, record in acdoca_records
        if normalize_key(record[acdoca_column_indices["gjahr"]]) == "2026"
        if normalize_account(record[acdoca_column_indices["racct"]]).startswith(
            "33"
        )
        and normalize_asset_code(record[acdoca_column_indices["anln1"]])
    }

    base_column_indices = populated_column_indices(base_headers, base_records)
    export_column_indices = populated_column_indices(export_headers, export_records)
    visible_base_headers = [base_headers[index] for index in base_column_indices]
    visible_export_headers = [export_headers[index] for index in export_column_indices]

    export_rows_by_equipment: dict[
        str, list[tuple[int, list[Any]]]
    ] = {}
    ih08_by_material_document: dict[
        tuple[str, str], list[tuple[int, list[Any]]]
    ] = {}
    for source_row, record in export_records_with_rows:
        key = normalize_key(record[EXPORT_EQUIPMENT_COLUMN - 1])
        if key:
            export_rows_by_equipment.setdefault(key, []).append(
                (source_row, record)
            )

        material_document = normalize_numeric_identifier(
            source_value(record, export_headers, "Documento material"), 10
        )
        material_year = normalize_numeric_identifier(
            source_value(record, export_headers, "Ejerc.documento mat."), 4
        )
        if material_document and material_year:
            ih08_by_material_document.setdefault(
                (material_document, material_year), []
            ).append((source_row, record))

    export_by_equipment = {
        equipment: rows[0][1]
        for equipment, rows in export_rows_by_equipment.items()
        if len(rows) == 1
    }
    export_row_by_equipment = {
        equipment: rows[0][0]
        for equipment, rows in export_rows_by_equipment.items()
        if len(rows) == 1
    }

    def ih08_asset_identity(
        rows: list[tuple[int, list[Any]]],
    ) -> tuple[str, str, bool]:
        asset_codes = {
            normalize_asset_code(
                source_value(record, export_headers, "Activo fijo")
            )
            for _, record in rows
        }
        asset_codes.discard("")
        if len(asset_codes) != 1:
            return (next(iter(asset_codes)) if len(asset_codes) == 1 else "", "", len(asset_codes) > 1)

        subnumbers = {
            normalize_asset_subnumber(
                first_populated(
                    source_value(record, export_headers, "Subnúmero activo fijo"),
                    source_value(record, export_headers, "Subnúmero"),
                )
            )
            for _, record in rows
        }
        subnumbers.discard("")
        if len(subnumbers) > 1:
            return next(iter(asset_codes)), "", True
        return next(iter(asset_codes)), next(iter(subnumbers), ""), False

    def validate_ar02_primary(
        rows: list[tuple[int, list[Any]]] | None,
    ) -> dict[str, Any]:
        if not rows:
            return {"status": "EQUIPO SIN VERSIÓN VIGENTE EN IH08"}
        if len(rows) > 1:
            return {"status": "REVISAR EQUIPO DUPLICADO"}

        asset_code, subnumber, ambiguous = ih08_asset_identity(rows)
        if ambiguous:
            return {"status": "REVISAR IDENTIDAD IH08", "anln1": asset_code}
        if not asset_code:
            return {"status": "EQUIPO SIN ACTIVO FIJO EN IH08"}

        asset_rows = ar02_by_asset.get(asset_code, [])
        if not asset_rows:
            return {
                "status": "NO EXISTE EN AR02",
                "anln1": asset_code,
                "anln2": subnumber,
            }
        if not subnumber:
            subnumbers = {
                normalize_asset_subnumber(
                    record[ar02_column_indices["anln2"]]
                )
                for _, record in asset_rows
            }
            if len(subnumbers) != 1:
                return {
                    "status": "REVISAR SUBNÚMERO",
                    "anln1": asset_code,
                }
            subnumber = next(iter(subnumbers))

        matches = ar02_by_key.get((asset_code, subnumber), [])
        if not matches:
            return {
                "status": "REVISAR SUBNÚMERO",
                "anln1": asset_code,
                "anln2": subnumber,
            }
        if len(matches) != 1:
            return {
                "status": "REVISAR DUPLICADO AR02",
                "anln1": asset_code,
                "anln2": subnumber,
            }

        # Control de centro: el activo de AR02 debe estar en el centro de
        # planificación del equipo en IH08. Si no, se busca un único activo
        # con el mismo nombre en ese centro; si no existe, no se asigna.
        ih08_record = rows[0][1]
        ih08_center = normalize_key(
            source_value(ih08_record, export_headers, "Centro planificación")
        )
        ar02_center = normalize_key(
            matches[0][1][ar02_column_indices["center"]]
        )
        route = "IH08 ACTIVO FIJO"
        fleet_in_ag09 = ih08_center.startswith("fl") and ar02_center == "ag09"
        if (
            ih08_center
            and ar02_center
            and ih08_center != ar02_center
            and not fleet_in_ag09
        ):
            ih08_name = normalize_search_text(
                source_value(ih08_record, export_headers, "Denominación objeto")
            )
            alternatives = ar02_by_name_center.get((ih08_name, ih08_center), [])
            if len(alternatives) != 1:
                return {
                    "status": "REVISAR CENTRO AR02 ≠ IH08",
                    "anln1": asset_code,
                    "anln2": subnumber,
                    "detail": (
                        f"IH08 trae {asset_code} (centro AR02 "
                        f"{ar02_center.upper()}), pero el equipo está en "
                        f"{ih08_center.upper()}."
                    ),
                }
            asset_code, subnumber = alternatives[0]
            matches = ar02_by_key[(asset_code, subnumber)]
            route = "NOMBRE IH08 + CENTRO IH08 EN AR02"
        return {
            "status": "VALIDADO AR02",
            "anln1": asset_code,
            "anln2": subnumber,
            "source_row": matches[0][0],
            "record": matches[0][1],
            "route": route,
        }

    def ar02_is_zero(record: list[Any]) -> bool:
        return all(
            (amount_key(record[index]) or Decimal(0)) == 0
            for index in ar02_value_indices
            if index is not None
        )

    def ih08_legacy_plate(rows: list[tuple[int, list[Any]]]) -> str:
        if ih08_plate_index is None:
            return ""
        values = {
            cell_text(record[ih08_plate_index])
            for _, record in rows
            if ih08_plate_index < len(record)
        }
        values.discard("")
        values.discard("-")
        return " / ".join(sorted(values))

    def find_legacy_entry(
        code: str, subnumber: str
    ) -> tuple[tuple[str, str] | None, dict[str, Any] | None]:
        if not code:
            return None, None
        keys = ar02_legacy_by_code.get(code, [])
        if subnumber and (code, subnumber) in ar02_legacy:
            key = (code, subnumber)
        elif (code, "0000") in ar02_legacy and not subnumber:
            key = (code, "0000")
        elif len(keys) == 1 and not subnumber:
            key = keys[0]
        else:
            return None, None
        return key, ar02_legacy[key]

    def hana_from_legacy(code: str, subnumber: str) -> tuple[list[tuple[str, str]], str]:
        """Activos HANA cuyo 'AF antiguo' en el AR02 es el código R3 dado."""
        keys = sorted(ar02_by_legacy.get(code, set()))
        if subnumber:
            keys = [key for key in keys if legacy_by_hana[key][1] == subnumber]
        elif len(keys) > 1:
            main = [key for key in keys if legacy_by_hana[key][1] == "0000"]
            keys = main or keys
        return keys, ""

    def apply_fallbacks(
        validation: dict[str, Any],
        rows: list[tuple[int, list[Any]]] | None,
        extra_plates: set[str] | frozenset[str] = frozenset(),
        equipment: str = "",
    ) -> dict[str, Any]:
        """Respaldo cuando IH08 no lleva a un activo válido en AR02, o cuando
        el activo HANA tiene todos sus valores en 0. Se prueba, en orden:

        1. Placa antigua de IH08 con forma de AF R3 → 'Activo fijo orig.' AR02.
        2. Activo fijo de IH08 que en realidad es un número R3.
        3. Placas (IH08 Número de inventario / placa antigua, placa de la
           BASE) → columnas 'Placa Activo Fijo' / 'Placa antigua' del AR02.

        Solo se cambia de activo si el candidato es único. Además se adjunta
        lo que dice el AR02 ANTIGUO del código R3 del activo final."""
        rows = rows or []
        plate_raw = ih08_legacy_plate(rows)
        plate_code, plate_sub = normalize_legacy_asset(plate_raw)
        plates = set(extra_plates)
        for _, record in rows:
            plates |= split_plates(
                source_value(record, export_headers, "Número de inventario")
            )
        if not plate_code:
            plates |= split_plates(plate_raw)
        if not equipment and rows:
            equipment = normalize_key(rows[0][1][EXPORT_EQUIPMENT_COLUMN - 1])

        primary_ok = validation["status"] == "VALIDADO AR02"
        primary_zero = primary_ok and ar02_is_zero(validation["record"])
        current_key = (
            (validation["anln1"], validation["anln2"]) if primary_ok else None
        )
        notes: list[str] = []
        if plate_raw and not plate_code and not split_plates(plate_raw):
            notes.append(
                f"Placa antigua '{plate_raw}' no es AF R3 ni placa reconocible."
            )

        ih08_r3 = ""
        if validation["status"] == "NO EXISTE EN AR02":
            candidate_code, _ = normalize_legacy_asset(
                validation.get("anln1", "").lstrip("0")
            )
            if candidate_code and (
                candidate_code in ar02_by_legacy
                or candidate_code in ar02_legacy_by_code
            ):
                ih08_r3 = candidate_code

        routes: list[tuple[str, list[tuple[str, str]]]] = []
        if plate_code:
            routes.append(
                ("IH08 PLACA ANTIGUA → AF R3 → AR02", hana_from_legacy(plate_code, plate_sub)[0])
            )
        if ih08_r3:
            routes.append(
                ("IH08 ACTIVO FIJO R3 → AF R3 → AR02", hana_from_legacy(ih08_r3, "")[0])
            )
        plate_keys = sorted(
            {key for plate in plates for key in ar02_by_plate.get(plate, ())}
        )
        if plate_keys:
            routes.append(
                (f"PLACA {' / '.join(sorted(plates))} → AR02", plate_keys)
            )

        if not primary_ok or primary_zero:
            for label, keys in routes:
                if len(keys) > 1:
                    notes.append(
                        f"{label}: varios activos HANA ("
                        + ", ".join(f"{a}-{b}" for a, b in keys)
                        + ")."
                    )
                    continue
                if not keys or keys[0] == current_key:
                    continue
                matches = ar02_by_key.get(keys[0], [])
                if len(matches) != 1:
                    continue
                if primary_ok and ar02_is_zero(matches[0][1]):
                    continue
                previous = validation["status"]
                if primary_ok:
                    previous = (
                        f"IH08 {current_key[0]}-{current_key[1]} con valores en 0"
                    )
                validation = {
                    "status": "VALIDADO AR02",
                    "anln1": keys[0][0],
                    "anln2": keys[0][1],
                    "source_row": matches[0][0],
                    "record": matches[0][1],
                    "route": label,
                    "detail": f"Antes: {previous}.",
                }
                primary_ok = True
                break
            else:
                if plate_code and not hana_from_legacy(plate_code, plate_sub)[0] and ar02_by_legacy:
                    notes.append(
                        f"AF R3 {plate_code} no figura como 'Activo fijo orig.' "
                        "en el AR02 HANA."
                    )

        # Código R3 para comparar con el AR02 ANTIGUO.
        mapped = (
            legacy_by_hana.get((validation["anln1"], validation["anln2"]))
            if primary_ok
            else None
        )
        if plate_code and mapped and mapped[0] != plate_code:
            notes.append(
                f"Placa antigua IH08 {plate_code} ≠ AF orig. del AR02 {mapped[0]}."
            )
        if plate_code:
            r3_code, r3_sub, r3_source = plate_code, plate_sub, "IH08 placa antigua"
        elif mapped:
            r3_code, r3_sub = mapped
            r3_source = "AR02 HANA: Activo fijo orig."
        elif ih08_r3:
            r3_code, r3_sub, r3_source = ih08_r3, "", "IH08 Activo fijo con número R3"
        else:
            r3_code, r3_sub, r3_source = "", "", ""
        legacy_key, legacy_entry = find_legacy_entry(
            r3_code,
            r3_sub or (validation.get("anln2", "") if primary_ok else ""),
        )
        validation["legacy"] = {
            "plate": plate_raw,
            "code": r3_code,
            "subnumber": r3_sub,
            "source": r3_source,
            "notes": notes,
            "key": legacy_key,
            "entry": legacy_entry,
        }
        return validation

    def validate_ar02(
        rows: list[tuple[int, list[Any]]] | None,
        extra_plates: set[str] | frozenset[str] = frozenset(),
    ) -> dict[str, Any]:
        return apply_fallbacks(validate_ar02_primary(rows), rows, extra_plates)

    def legacy_output(validation: dict[str, Any]) -> list[Any]:
        legacy = validation.get("legacy") or {}
        code = legacy.get("code", "")
        entry = legacy.get("entry")
        key = legacy.get("key")
        source = " | ".join(
            part for part in [legacy.get("source", ""), *legacy.get("notes", [])]
            if part
        )
        route = validation.get("route", "")
        if not code:
            return [route, legacy.get("plate", ""), "", source, *[""] * 9]
        shown_code = f"{code}-{key[1]}" if key else code
        if entry is None:
            state = (
                "NO EXISTE EN AR02 ANTIGUO"
                if ar02_legacy
                else "SIN ARCHIVO AR02 ANTIGUO"
            )
            return [route, legacy.get("plate", ""), shown_code, source, state, *[""] * 8]
        record = validation.get("record")
        if validation["status"] != "VALIDADO AR02" or record is None:
            comparison = "SOLO EN AR02 ANTIGUO"
        else:
            def hana_amount(header: str) -> Decimal:
                index = ar02_header_indices[normalize_header(header)]
                return amount_key(record[index]) or Decimal(0)

            cap = hana_amount("CAP actuales")
            book = hana_amount("Val.cont.act.")
            if ar02_is_zero(record) and (entry["cap"] != 0 or entry["valcont"] != 0):
                comparison = "REVISAR: HANA EN 0, ANTIGUO CON VALOR"
            elif abs(cap - entry["cap"]) <= Decimal("0.01") and abs(
                book - entry["valcont"]
            ) <= Decimal("0.01"):
                comparison = "COINCIDE"
            else:
                comparison = "DIFERENTE"
        return [
            route,
            legacy.get("plate", ""),
            shown_code,
            source,
            "ENCONTRADO"
            + (f" ({len(entry['rows'])} filas sumadas)" if len(entry["rows"]) > 1 else ""),
            entry["denom"],
            float(entry["cap"]),
            float(entry["baja"]),
            float(entry["amo"]),
            float(entry["valcont"]),
            legacy_asset_state(entry),
            comparison,
            " / ".join(str(row) for row in entry["rows"]),
        ]

    acdoca_documents_by_asset: dict[tuple[str, str], dict[str, Any]] = {}
    for asset_key, movements in acdoca_by_asset_key.items():
        document_keys = sorted(
            {
                (
                    normalize_key(record[acdoca_column_indices["gjahr"]]),
                    normalize_numeric_identifier(
                        record[acdoca_column_indices["belnr"]], 10
                    ),
                )
                for _, record in movements
                if normalize_numeric_identifier(
                    record[acdoca_column_indices["belnr"]], 10
                )
            }
        )
        documents = [
            f"{document} ({year})" for year, document in document_keys
        ]
        accounts = sorted(
            {
                normalize_account(record[acdoca_column_indices["racct"]])
                for _, record in movements
                if normalize_account(
                    record[acdoca_column_indices["racct"]]
                ).startswith(("32", "33"))
            }
        )
        acdoca_documents_by_asset[asset_key] = {
            "documents": documents,
            "accounts": accounts,
            "movements": movements,
        }

    bkpf_by_awkey: dict[str, list[tuple[str, int, list[Any]]]] = {}
    bkpf_by_any_awkey: dict[str, list[tuple[str, int, list[Any]]]] = {}
    for source_file, source_row, record in bkpf_records:
        awkey = normalize_key(record[bkpf_column_indices["awkey"]])
        if not awkey:
            continue
        entry = (source_file, source_row, record)
        bkpf_by_any_awkey.setdefault(awkey, []).append(entry)
        if normalize_key(record[bkpf_column_indices["awtyp"]]) == "mkpf":
            bkpf_by_awkey.setdefault(awkey, []).append(entry)

    base_equipment_counts = Counter(
        normalize_key(record[BASE_EQUIPMENT_COLUMN - 1])
        for record in base_records
        if len(record) >= BASE_EQUIPMENT_COLUMN
        and normalize_key(record[BASE_EQUIPMENT_COLUMN - 1])
    )
    audit_headers = [
        "ID BASE",
        "FILA BASE ORIGEN",
        "EQUNR",
        "FILA IH08 ORIGEN",
        "ANLN1 IH08",
        "RUTA DE CONCILIACIÓN",
        "ESTADO",
        "MOTIVO / EVIDENCIA",
        "DOCUMENTO MATERIAL IH08",
        "EJERCICIO DOCUMENTO MATERIAL",
        "FECHA IH08",
        "AWKEY ESPERADO",
        "AWTYP BKPF",
        "AWKEY BKPF",
        "BUKRS BKPF",
        "BELNR BKPF",
        "GJAHR BKPF",
        "BUDAT BKPF",
        "FILA BKPF ORIGEN",
        "ARCHIVO BKPF ORIGEN",
        "RBUKRS ACDOCA",
        "GJAHR ACDOCA",
        "BELNR ACDOCA",
        "RLDNR ACDOCA",
        "POSICIÓN ACDOCA",
        "RACCT ACDOCA",
        "HSL ACDOCA",
        "BUDAT ACDOCA",
        "ANLN1 ACDOCA",
        "ANLN2 ACDOCA",
        "FILA ACDOCA ORIGEN",
        "ARCHIVO ACDOCA ORIGEN",
    ]
    bkpf_audit_records: list[list[Any]] = []

    def add_audit_rows(
        base_record: list[Any],
        base_source_row: int,
        export_record: list[Any],
        route: str,
        status: str,
        reason: str,
        export_source_row: int | None = None,
        expected_awkey: str | None = None,
        bkpf_entry: tuple[str, int, list[Any]] | None = None,
        acdoca_entry: tuple[int, list[Any]] | None = None,
    ) -> None:
        bkpf_record = bkpf_entry[2] if bkpf_entry else []
        acdoca_record = acdoca_entry[1] if acdoca_entry else []
        get_bkpf = lambda field: (
            bkpf_record[bkpf_column_indices[field]] if bkpf_record else None
        )
        get_acdoca = lambda field: (
            acdoca_record[acdoca_column_indices[field]]
            if acdoca_record
            else None
        )
        bkpf_audit_records.append(
            [
                source_value(base_record, base_headers, "ID"),
                base_source_row,
                source_value(base_record, base_headers, "Código de equipo (si aplica)"),
                export_source_row,
                source_value(export_record, export_headers, "Activo fijo"),
                route,
                status,
                reason,
                source_value(export_record, export_headers, "Documento material"),
                source_value(export_record, export_headers, "Ejerc.documento mat."),
                source_value(export_record, export_headers, "Fecha contabiliz."),
                expected_awkey,
                get_bkpf("awtyp"),
                get_bkpf("awkey"),
                get_bkpf("bukrs"),
                get_bkpf("belnr"),
                get_bkpf("gjahr"),
                get_bkpf("budat"),
                bkpf_entry[1] if bkpf_entry else None,
                bkpf_entry[0] if bkpf_entry else None,
                get_acdoca("rbukrs"),
                get_acdoca("gjahr"),
                get_acdoca("belnr"),
                get_acdoca("rldnr"),
                get_acdoca("docln"),
                get_acdoca("racct"),
                get_acdoca("hsl"),
                get_acdoca("budat"),
                get_acdoca("anln1"),
                get_acdoca("anln2"),
                acdoca_entry[0] if acdoca_entry else None,
                acdoca_path.name if acdoca_entry else None,
            ]
        )

    def output_tuple(record: list[Any]) -> tuple[Any, ...]:
        try:
            amount = amount_key(record[acdoca_column_indices["hsl"]])
        except ValueError:
            raise
        return (
            normalize_numeric_identifier(
                record[acdoca_column_indices["belnr"]], 10
            ),
            normalize_key(record[acdoca_column_indices["racct"]]),
            amount,
            date_key(record[acdoca_column_indices["budat"]]),
        )

    def output_values(record: list[Any]) -> list[Any]:
        return [
            record[acdoca_column_indices[field]]
            for field in ("belnr", "racct", "hsl", "budat")
        ]

    def reconcile_record(
        base_record: list[Any],
        base_source_row: int,
        export_record: list[Any] | None,
        export_source_row: int | None,
        validation: dict[str, Any],
    ) -> tuple[
        list[Any],
        str,
        int | None,
        list[tuple[int, list[Any]]],
        str,
    ]:
        equipment = normalize_key(
            base_record[BASE_EQUIPMENT_COLUMN - 1]
            if len(base_record) >= BASE_EQUIPMENT_COLUMN
            else None
        )
        if validation["status"] != "VALIDADO AR02":
            status = validation["status"]
            add_audit_rows(
                base_record,
                base_source_row,
                export_record or [],
                "IH08 → AR02",
                status,
                "No se conciliaron documentos porque la identidad del activo "
                "no quedó validada de forma única en AR02.",
                export_source_row,
            )
            return [None] * 4, status, None, [], "IH08 → AR02"

        hana_code = validation["anln1"]
        hana_subnumber = validation["anln2"]
        asset_movements = acdoca_by_asset_2026.get(
            (hana_code, hana_subnumber), []
        )
        if len(asset_movements) == 1:
            source_row, record = asset_movements[0]
            status = "Coincidencia única"
            add_audit_rows(
                base_record,
                base_source_row,
                export_record or [],
                "ANLN1 + ANLN2",
                status,
                "Activo de IH08 validado en AR02; ANLN1 y ANLN2 coinciden "
                "con ACDOCA.",
                export_source_row,
                acdoca_entry=asset_movements[0],
            )
            return (
                output_values(record),
                status,
                source_row,
                [],
                "ANLN1 + ANLN2",
            )
        if len(asset_movements) > 1:
            status = f"Múltiples movimientos ({len(asset_movements)}) - revisar"
            for movement in asset_movements:
                add_audit_rows(
                    base_record,
                    base_source_row,
                    export_record or [],
                    "ANLN1 + ANLN2",
                    status,
                    "ANLN1 y ANLN2 coinciden con varias posiciones; no se "
                    "eligió una.",
                    export_source_row,
                    acdoca_entry=movement,
                )
            return (
                [None] * 4,
                status,
                None,
                asset_movements,
                "ANLN1 + ANLN2",
            )

        if not export_record:
            status = "AR02 validado; IH08 ambiguo"
            add_audit_rows(
                base_record,
                base_source_row,
                [],
                "IH08 ambiguo",
                status,
                "La identidad del activo coincide en las filas IH08, pero "
                "no se eligió una fila de cabecera entre varias.",
            )
            return [None] * 4, status, None, [], "IH08 ambiguo"

        material_document = normalize_numeric_identifier(
            source_value(export_record, export_headers, "Documento material"), 10
        )
        material_year = normalize_numeric_identifier(
            source_value(export_record, export_headers, "Ejerc.documento mat."), 4
        )
        if not material_document or not material_year:
            status = "Pendiente: sin documento material"
            add_audit_rows(
                base_record,
                base_source_row,
                export_record,
                "BKPF/MKPF → ACDOCA",
                status,
                "IH08 no contiene MBLNR y ejercicio válidos.",
                export_source_row,
            )
            return [None] * 4, status, None, [], "BKPF/MKPF → ACDOCA"

        expected_awkey = material_document + material_year
        mkpf_entries = bkpf_by_awkey.get(expected_awkey, [])
        if not mkpf_entries:
            other_entries = bkpf_by_any_awkey.get(expected_awkey, [])
            status = (
                "Pendiente: AWTYP no es MKPF"
                if other_entries
                else "Pendiente: sin BKPF MKPF/AWKEY"
            )
            reason = (
                "AWKEY existe en BKPF, pero AWTYP no es MKPF."
                if other_entries
                else "No existe BKPF con AWTYP=MKPF y AWKEY exacto."
            )
            if other_entries:
                for entry in other_entries:
                    add_audit_rows(
                        base_record,
                        base_source_row,
                        export_record,
                        "BKPF/MKPF → ACDOCA",
                        status,
                        reason,
                        export_source_row,
                        expected_awkey,
                        bkpf_entry=entry,
                    )
            else:
                add_audit_rows(
                    base_record,
                    base_source_row,
                    export_record,
                    "BKPF/MKPF → ACDOCA",
                    status,
                    reason,
                    export_source_row,
                    expected_awkey,
                )
            return [None] * 4, status, None, [], "BKPF/MKPF → ACDOCA"

        if len(mkpf_entries) != 1:
            status = "Pendiente: AWKEY ambiguo en BKPF"
            for entry in mkpf_entries:
                add_audit_rows(
                    base_record,
                    base_source_row,
                    export_record,
                    "BKPF/MKPF → ACDOCA",
                    status,
                    f"AWKEY exacto coincide con {len(mkpf_entries)} documentos BKPF.",
                    export_source_row,
                    expected_awkey,
                    bkpf_entry=entry,
                )
            return [None] * 4, status, None, [], "BKPF/MKPF → ACDOCA"

        bkpf_entry = mkpf_entries[0]
        bkpf_record = bkpf_entry[2]
        bukrs = normalize_key(bkpf_record[bkpf_column_indices["bukrs"]])
        gjahr = normalize_key(bkpf_record[bkpf_column_indices["gjahr"]])
        belnr = normalize_numeric_identifier(
            bkpf_record[bkpf_column_indices["belnr"]], 10
        )
        accounting_document = acdoca_by_document.get(
            (bukrs, gjahr, belnr), []
        )
        reasons: list[str] = []
        if bukrs != "ag10" or gjahr != material_year or gjahr != "2026":
            reasons.append("Sociedad o ejercicio BKPF no compatible.")
        if not accounting_document:
            reasons.append("No se encontró BELNR/BUKRS/GJAHR en ACDOCA.")
        if accounting_document and any(
            normalize_key(record[acdoca_column_indices["rldnr"]]) != "0l"
            for _, record in accounting_document
        ):
            reasons.append("Ledger ACDOCA distinto de 0L.")
        ih08_date = date_key(
            source_value(export_record, export_headers, "Fecha contabiliz.")
        )
        bkpf_date = date_key(bkpf_record[bkpf_column_indices["budat"]])
        acdoca_dates = {
            date_key(record[acdoca_column_indices["budat"]])
            for _, record in accounting_document
        }
        if (
            ih08_date is None
            or bkpf_date is None
            or not acdoca_dates
            or acdoca_dates != {bkpf_date}
            or ih08_date != bkpf_date
        ):
            reasons.append("Fechas IH08/BKPF/ACDOCA ausentes o inconsistentes.")

        accounting_tuples = {
            output_tuple(record) for _, record in accounting_document
        }
        if accounting_document and len(accounting_tuples) != 1:
            reasons.append(
                "Las posiciones ACDOCA tienen valores contables distintos."
            )

        material_key = (material_document, material_year)
        related_ih08_rows = ih08_by_material_document.get(material_key, [])
        if len(related_ih08_rows) != 1:
            reasons.append(
                "El documento material aparece en varias filas/equipos de IH08."
            )
        if equipment and base_equipment_counts[equipment] > 1:
            reasons.append("El mismo EQUNR aparece en varios registros BASE.")

        status = (
            "Coincidencia BKPF/ACDOCA"
            if not reasons
            else "Pendiente: " + " ".join(reasons)
        )
        if accounting_document:
            for movement in accounting_document:
                add_audit_rows(
                    base_record,
                    base_source_row,
                    export_record,
                    "EQUNR → IH08 MBLNR/MJAHR → BKPF MKPF/AWKEY → ACDOCA BELNR",
                    status,
                    " ".join(reasons)
                    if reasons
                    else "AWKEY, documento contable, sociedad, ejercicio y fechas exactos; "
                    "se conservan todas las posiciones ACDOCA.",
                    export_source_row,
                    expected_awkey,
                    bkpf_entry,
                    movement,
                )
        else:
            add_audit_rows(
                base_record,
                base_source_row,
                export_record,
                "EQUNR → IH08 MBLNR/MJAHR → BKPF MKPF/AWKEY → ACDOCA BELNR",
                status,
                " ".join(reasons),
                export_source_row,
                expected_awkey,
                bkpf_entry,
            )

        if reasons:
            return [None] * 4, status, None, [], "BKPF/MKPF → ACDOCA"
        representative = accounting_document[0][1]
        return (
            output_values(representative),
            status,
            None,
            [],
            "BKPF/MKPF → ACDOCA",
        )

    merged_headers = visible_base_headers + [
        f"IH08 - {header}" for header in visible_export_headers
    ]
    merged_headers.append("Resultado del cruce")
    merged_headers.extend(ACDOCA_OUTPUT_HEADERS)
    merged_headers.extend(AR02_OUTPUT_HEADERS)
    merged_records: list[list[Any]] = []
    unmatched_records: list[list[Any]] = []
    acdoca_review_records: list[list[Any]] = []
    validations_by_base_row: dict[int, dict[str, Any]] = {}
    ar02_output_by_base_row: dict[int, list[Any]] = {}
    reconciliations_by_base_row: dict[
        int, tuple[list[Any], str, int | None, list[tuple[int, list[Any]]], str]
    ] = {}
    matched_count = 0
    recovered_base_rows: list[list[Any]] = []

    for base_record, base_source_row in zip(base_records, base_source_rows):
        key = normalize_key(
            base_record[BASE_EQUIPMENT_COLUMN - 1]
            if len(base_record) >= BASE_EQUIPMENT_COLUMN
            else None
        )
        export_record = export_by_equipment.get(key) if key else None
        export_rows = export_rows_by_equipment.get(key, []) if key else []
        export_source_row = export_row_by_equipment.get(key)
        base_plates = split_plates(
            source_value(base_record, base_headers, "Código placa (si aplica)")
        )
        if not key:
            validation = apply_fallbacks(
                {"status": "BASE SIN CÓDIGO DE EQUIPO"}, None, base_plates
            )
        elif key not in ih08_equipment_seen:
            validation = apply_fallbacks(
                {"status": "EQUIPO NO EXISTE EN IH08"}, None, base_plates, key
            )
        else:
            validation = validate_ar02(export_rows or None, base_plates)
        validations_by_base_row[base_source_row] = validation
        reconciliation = reconcile_record(
            base_record,
            base_source_row,
            export_record,
            export_source_row,
            validation,
        )
        reconciliations_by_base_row[base_source_row] = reconciliation
        accounting_values, accounting_status, accounting_source_row, movements, _ = (
            reconciliation
        )
        accounting_output = (
            accounting_values
            + [accounting_status, base_source_row, accounting_source_row]
        )
        asset_key = (
            (validation["anln1"], validation["anln2"])
            if validation["status"] == "VALIDADO AR02"
            else None
        )
        asset_data = (
            acdoca_documents_by_asset.get(asset_key, {})
            if asset_key is not None
            else {}
        )
        documents = asset_data.get("documents", [])
        accounts = asset_data.get("accounts", [])
        ar02_record = validation.get("record")
        if ar02_record is not None:
            cap_account = normalize_account(
                ar02_record[ar02_column_indices["cta_cap"]]
            )
            account_control = (
                ""
                if not accounts
                else (
                    "REVISAR CUENTA"
                    if any(account != cap_account for account in accounts)
                    else "OK"
                )
            )
            ar02_output = [
                validation["status"],
                (
                    "CON MOVIMIENTOS"
                    if documents
                    else "SIN MOVIMIENTOS EN ACDOCA"
                ),
                f'{validation["anln1"]}-{validation["anln2"]}',
                ar02_record[ar02_column_indices["denom"]],
                cap_account,
                ar02_record[ar02_column_indices["clase"]],
                ar02_record[ar02_column_indices["capitalized"]],
                *[
                    ar02_record[index] if index is not None else None
                    for index in ar02_value_indices
                ],
                ar02_asset_state(ar02_record, ar02_header_indices),
                " / ".join(documents),
                len(documents),
                " / ".join(accounts),
                account_control,
                validation["source_row"],
            ]
        else:
            ar02_output = [
                validation["status"],
                "",
                f'{validation["anln1"]}-{validation["anln2"]}'
                if validation.get("anln1")
                else "",
                "",
                "",
                "",
                "",
                *([None] * len(AR02_VALUE_HEADERS)),
                "",
                "",
                0,
                "",
                "",
                "",
            ]
        ar02_output.extend(legacy_output(validation))
        if len(ar02_output) != len(AR02_OUTPUT_HEADERS):
            raise ValueError(
                f"Fila BASE {base_source_row}: la salida AR02 tiene "
                f"{len(ar02_output)} columnas y se esperaban "
                f"{len(AR02_OUTPUT_HEADERS)}."
            )
        ar02_output_by_base_row[base_source_row] = ar02_output
        accounting_output.extend(ar02_output)

        ih08_asset_code, _, _ = ih08_asset_identity(export_rows)
        if (
            validation["status"] == "VALIDADO AR02"
            and ih08_asset_code
            and ih08_asset_code not in baseline_acdoca_asset_codes
        ):
            recovered_base_rows.append(
                [
                    source_value(base_record, base_headers, "ID"),
                    base_source_row,
                    key,
                    validation["anln1"],
                    validation["anln2"],
                    ar02_record[ar02_column_indices["denom"]],
                    validation["status"],
                ]
            )
        if len(movements) > 1:
            for acdoca_source_row, acdoca_record in movements:
                acdoca_review_records.append(
                    [
                        source_value(base_record, base_headers, "ID"),
                        base_source_row,
                        source_value(export_record or [], export_headers, "Activo fijo"),
                        acdoca_source_row,
                        *acdoca_record,
                    ]
                )

        if export_record is None:
            status = (
                "Pendiente: IH08 ambiguo"
                if export_rows
                else "Sin coincidencia"
            )
            unmatched_records.append(
                [base_record[index] for index in base_column_indices]
                + [status]
                + accounting_output
            )
            merged_records.append(
                [base_record[index] for index in base_column_indices]
                + [None] * len(export_column_indices)
                + [status]
                + accounting_output
            )
            continue

        matched_count += 1
        material = export_record[EXPORT_MATERIAL_COLUMN - 1]
        serial = export_record[EXPORT_SERIAL_COLUMN - 1]
        if material is not None and cell_text(material):
            base_record[BASE_MATERIAL_COLUMN - 1] = cell_text(material)
        if serial is not None and cell_text(serial):
            base_record[BASE_SERIAL_COLUMN - 1] = cell_text(serial)
        merged_records.append(
            [base_record[index] for index in base_column_indices]
            + [export_record[index] for index in export_column_indices]
            + ["Coincide"]
            + accounting_output
        )

    # Documentos de baja por activo (ANLN1): clases de movimiento de baja,
    # documento + fecha de contabilización. Se excluye la clase de doc. AF.
    baja_documents_by_asset: dict[str, set[tuple[str, datetime]]] = {}
    for _, record in acdoca_records:
        movement = cell_text(record[acdoca_column_indices["anbwa"]])
        if not BAJA_MOVEMENT_PATTERN.fullmatch(movement):
            continue
        document_class_index = acdoca_column_indices.get("blart")
        if (
            document_class_index is not None
            and normalize_key(record[document_class_index]).upper() == "AF"
        ):
            continue
        asset_code = normalize_asset_code(record[acdoca_column_indices["anln1"]])
        document = cell_text(record[acdoca_column_indices["belnr"]]).lstrip("0")
        posting_date = record[acdoca_column_indices["budat"]]
        if asset_code and document.isdigit() and isinstance(posting_date, datetime):
            baja_documents_by_asset.setdefault(asset_code, set()).add(
                (document, posting_date)
            )

    def resolve_equipment(equipment: str) -> dict[str, Any]:
        if equipment not in ih08_equipment_seen:
            return {"status": "EQUIPO NO EXISTE EN IH08"}
        return validate_ar02(export_rows_by_equipment.get(equipment) or None)

    def ih08_fields(equipment: str) -> dict[str, Any] | None:
        record = export_by_equipment.get(equipment)
        if record is None:
            return None
        return {
            "plate": source_value(record, export_headers, "Número de inventario"),
            "manufacturer": source_value(record, export_headers, "Fabricante"),
            "material": source_value(record, export_headers, "Material"),
            "center": source_value(record, export_headers, "Centro planificación"),
        }

    def ar02_amount(record: list[Any], header: str) -> Decimal:
        return amount_key(record[ar02_header_indices[normalize_header(header)]]) or Decimal(0)

    def legacy_book_info(
        asset_code: str, validation: dict[str, Any] | None = None
    ) -> dict[str, Any] | None:
        """Datos del AR02 ANTIGUO para un ANLN1 HANA (todas sus subnúmeros).

        alert=True cuando el activo HANA está en 0 y el antiguo tiene valor."""
        entries: dict[tuple[str, str], dict[str, Any]] = {}
        legacy = (validation or {}).get("legacy") or {}
        if (
            validation
            and validation.get("anln1") == asset_code
            and legacy.get("entry") is not None
        ):
            entries[legacy["key"]] = legacy["entry"]
        for (hana_code, hana_sub), (code, sub) in legacy_by_hana.items():
            if hana_code != asset_code:
                continue
            key, entry = find_legacy_entry(code, sub)
            if entry is not None:
                entries[key] = entry
        rows = ar02_by_asset.get(asset_code, [])
        if not entries or not rows:
            return None
        hana_zero = all(ar02_is_zero(record) for _, record in rows)
        disposal = sum(entry["baja"] for entry in entries.values())
        if disposal != 0:
            value = -(disposal + sum(e["amo_baja"] for e in entries.values()))
            source = "AR02 ANTIGUO: valor neto dado de baja = -(Baja + Amo baja)"
        else:
            value = sum(entry["valcont"] for entry in entries.values())
            source = "AR02 ANTIGUO: Val.cont.act."
        codes = ", ".join(f"{code}-{sub}" for code, sub in sorted(entries))
        has_value = value != 0
        return {
            "alert": hana_zero and has_value,
            "value": value,
            "source": f"{source} (R3 {codes})",
            "detail": (
                f"HANA {asset_code} tiene CAP/Val.cont. en 0. AR02 ANTIGUO "
                f"R3 {codes}: CAP actuales "
                f"{sum(e['cap'] for e in entries.values())}, Val.cont.act. "
                f"{sum(e['valcont'] for e in entries.values())}, "
                f"valor sugerido {value}."
            ),
        }

    def r3_legacy_detail(r3_code: str) -> str:
        code, sub = normalize_legacy_asset(r3_code)
        _, entry = find_legacy_entry(code, sub)
        if entry is None:
            return "Tampoco está en el AR02 ANTIGUO."
        state = legacy_asset_state(entry)
        return (
            f"AR02 ANTIGUO: '{entry['denom']}', CAP actuales {entry['cap']}, "
            f"Val.cont.act. {entry['valcont']}"
            + (f" ({state})" if state else "")
            + "."
        )

    def hana_codes_from_r3(r3_code: str) -> list[str]:
        code, _ = normalize_legacy_asset(r3_code)
        return sorted({hana for hana, _ in ar02_by_legacy.get(code, set())})

    def asset_book_value(
        asset_code: str, validation: dict[str, Any] | None = None
    ) -> tuple[Decimal | None, str]:
        """Valor contable para el STATUS (suma de subnúmeros del ANLN1).

        Si el activo tiene BAJA en AR02 se usa el valor neto dado de baja
        (-(Baja + Amo baja)); si no, el valor contable actual.
        """
        rows = ar02_by_asset.get(asset_code, [])
        if not rows:
            return None, ""
        if USE_LEGACY_BOOK_VALUE_WHEN_HANA_IS_ZERO:
            legacy_info = legacy_book_info(asset_code, validation)
            if legacy_info and legacy_info["alert"]:
                return legacy_info["value"], legacy_info["source"]
        disposal = sum(ar02_amount(record, "Baja") for _, record in rows)
        disposal_depreciation = sum(
            ar02_amount(record, "Amo baja") for _, record in rows
        )
        if disposal != 0:
            return (
                -(disposal + disposal_depreciation),
                "AR02: valor neto dado de baja = -(Baja + Amo baja)",
            )
        return (
            sum(ar02_amount(record, "Val.cont.act.") for _, record in rows),
            "AR02: Val.cont.act.",
        )

    def asset_denomination(asset_code: str) -> str:
        rows = ar02_by_asset.get(asset_code, [])
        return cell_text(rows[0][1][ar02_column_indices["denom"]]) if rows else "no está en AR02"

    status_records: list[list[Any]] = []
    status_candidates: list[dict[str, Any]] = []
    current_month = SPANISH_MONTHS[datetime.now().month - 1]
    for base_record, base_source_row in zip(base_records, base_source_rows):
        key = normalize_key(source_value(base_record, base_headers, "Código de equipo (si aplica)"))
        export_record = export_by_equipment.get(key) if key else None
        export_record = export_record or []
        accounting_values, accounting_status, accounting_source_row, _, _ = (
            reconciliations_by_base_row[base_source_row]
        )
        get_base = lambda name: source_value(base_record, base_headers, name)
        get_export = lambda name: source_value(export_record, export_headers, name)
        provenance = cell_text(
            get_base("Procedencia del Equipo (Planta/ Embarcación)")
        )
        provenance_center = re.match(r"\s*((?:AG|FL)\s*\d{2})", provenance, re.I)
        base_id = cell_text(get_base("ID"))
        base_plate = normalize_status_plate(get_base("Código placa (si aplica)"))
        if base_plate in {"", "-"}:
            base_plate = ""
        ih08_plate = get_export("Número de inventario")
        # Columna del STATUS -> (valor, fuente). Patrón tomado de las filas
        # FIC-xxxxxx que ya están en el STATUS.
        values: dict[str, tuple[Any, str]] = {
            "EMISOR": (get_base("Solicitante"), "Solicitante"),
            "DOC. REFERENCIA": (
                f"FIC-{int(base_id):06d}" if base_id.isdigit() else base_id,
                "ID de la ficha",
            ),
            "DESCRIPCION DEL ACTIVO": (
                first_populated(
                    get_base("Descripción del equipo"),
                    get_export("Descripción material"),
                ),
                "Descripción del equipo",
            ),
            "PLACA": (
                base_plate
                or (
                    normalize_status_plate(ih08_plate)
                    if is_plate_like(ih08_plate)
                    else None
                ),
                "Código placa" if base_plate else "IH08 Número de inventario",
            ),
            "MARCA": (get_export("Fabricante"), "IH08 Fabricante"),
            "Nº SERIE": (
                first_populated(
                    get_export("Material"), get_base("Código de material")
                ),
                "IH08 Material / Código de material",
            ),
            "SERIE SAP": (
                first_populated(
                    get_export("Equipo"),
                    get_base("Código de equipo (si aplica)"),
                ),
                "Código de equipo",
            ),
            "CENTRO": (
                first_populated(
                    get_export("Centro planificación"),
                    provenance_center.group(1).replace(" ", "").upper()
                    if provenance_center
                    else None,
                ),
                "IH08 Centro planificación / Procedencia",
            ),
            "E/P": (
                provenance.split("-", 1)[1].strip()
                if provenance.upper().startswith("FL") and "-" in provenance
                else None,
                "Procedencia (embarcación)",
            ),
        }
        validation = validations_by_base_row[base_source_row]
        status_candidates.append(
            {
                "base_row": base_source_row,
                "equipment": key,
                "fic": fic_number(base_id),
                "description": get_base("Descripción del equipo"),
                "plate": base_plate,
                "values": values,
                "validation": validation,
            }
        )
        row_by_header = {header: value for header, (value, _) in values.items()}
        row_by_header["MES TRABAJADO"] = current_month
        row_by_header["MES RECEPCION FICHA DE ESTADO"] = current_month.capitalize()
        if validation["status"] == "VALIDADO AR02":
            row_by_header["CODIGO AF HANAH"] = validation["anln1"]
        status_records.append(
            [row_by_header.get(header) for header in STATUS_HEADERS]
            + [
                *accounting_values,
                accounting_status,
                base_source_row,
                accounting_source_row,
                *ar02_output_by_base_row[base_source_row],
            ]
        )

    candidate_headers = [
        "ID BASE",
        "DESCRIPCIÓN BASE",
        "ACTIVO SUGERIDO",
        "DENOMINACIÓN AR02",
        "PUNTAJE",
        "CENTRO",
    ]
    ar02_candidates = []
    for (asset_code, subnumber), rows in ar02_by_key.items():
        if len(rows) != 1:
            continue
        ar02_row = rows[0][1]
        center = cell_text(ar02_row[ar02_column_indices["center"]])
        denomination = cell_text(ar02_row[ar02_column_indices["denom"]])
        normalized_denomination = normalize_search_text(denomination)
        if not normalized_denomination:
            continue
        ar02_candidates.append(
            {
                "asset": f"{asset_code}-{subnumber}",
                "denomination": denomination,
                "center": center,
                "normalized": normalized_denomination,
            }
        )

    candidates_by_center: dict[str, list[dict[str, str]]] = {}
    for candidate in ar02_candidates:
        candidates_by_center.setdefault(
            normalize_key(candidate["center"]), []
        ).append(candidate)
    all_candidate_names = [candidate["normalized"] for candidate in ar02_candidates]
    candidate_records: list[list[Any]] = []
    candidate_examples: list[list[Any]] = []
    candidate_top_scores: dict[int, int] = {}
    candidate_review_statuses = {
        "EQUIPO SIN ACTIVO FIJO EN IH08",
        "BASE SIN CÓDIGO DE EQUIPO",
    }

    for base_record, base_source_row in zip(base_records, base_source_rows):
        if (
            validations_by_base_row[base_source_row]["status"]
            not in candidate_review_statuses
        ):
            continue

        base_description = source_value(
            base_record, base_headers, "Descripción del equipo"
        )
        query = normalize_search_text(base_description)
        provenance = cell_text(
            source_value(
                base_record,
                base_headers,
                "Procedencia del Equipo (Planta/ Embarcación)",
            )
        )
        technical_location = cell_text(
            source_value(
                base_record, base_headers, "Ubicación técnica del equipo"
            )
        )
        center_pattern = r"\b(?:AG|FL)\s*\d{2}\b"
        center_match = re.search(
            center_pattern, provenance, re.IGNORECASE
        ) or re.search(center_pattern, technical_location, re.IGNORECASE)
        center_code = (
            normalize_key(center_match.group(0)).replace(" ", "")
            if center_match
            else ""
        )
        filtered_candidates = candidates_by_center.get(center_code)
        if filtered_candidates:
            search_candidates = filtered_candidates
            choices = [
                candidate["normalized"] for candidate in search_candidates
            ]
        else:
            search_candidates = ar02_candidates
            choices = all_candidate_names

        matches = (
            process.extract(
                query,
                choices,
                scorer=fuzz.WRatio,
                limit=3,
            )
            if query and choices
            else []
        )
        if not matches:
            candidate_records.append(
                [
                    source_value(base_record, base_headers, "ID"),
                    cell_text(base_description),
                    None,
                    None,
                    None,
                    None,
                ]
            )
            continue

        candidate_top_scores[base_source_row] = int(matches[0][1])
        first_candidate_record: list[Any] | None = None
        for _, score, candidate_index in matches:
            candidate = search_candidates[candidate_index]
            suggestion = [
                source_value(base_record, base_headers, "ID"),
                cell_text(base_description),
                candidate["asset"],
                candidate["denomination"],
                round(score, 1),
                candidate["center"],
            ]
            candidate_records.append(suggestion)
            if first_candidate_record is None:
                first_candidate_record = suggestion
        if first_candidate_record is not None:
            candidate_examples.append(first_candidate_record)

    candidate_high_score_count = sum(
        score >= 90 for score in candidate_top_scores.values()
    )
    ar02_sheet_headers = ar02_headers + [
        "ANLN1 NORMALIZADO",
        "ANLN2 NORMALIZADO",
        "CLAVE AR02",
        "N_DOCUMENTOS",
        "DOCUMENTOS ACDOCA",
        "CUENTAS 32/33 ACDOCA",
        "ESTADO MOVIMIENTOS",
        "CONTROL CUENTA",
        "FILA AR02 ORIGEN",
    ]
    ar02_sheet_records: list[list[Any]] = []
    for (asset_code, subnumber), rows in sorted(ar02_by_key.items()):
        summary = acdoca_documents_by_asset.get(
            (asset_code, subnumber),
            {"documents": [], "accounts": []},
        )
        for source_row, record in rows:
            cap_account = normalize_account(
                record[ar02_column_indices["cta_cap"]]
            )
            accounts = summary["accounts"]
            account_control = (
                ""
                if not accounts
                else (
                    "REVISAR CUENTA"
                    if any(account != cap_account for account in accounts)
                    else "OK"
                )
            )
            ar02_sheet_records.append(
                record
                + [
                    asset_code,
                    subnumber,
                    f"{asset_code}-{subnumber}",
                    len(summary["documents"]),
                    " / ".join(summary["documents"]),
                    " / ".join(accounts),
                    (
                        "CON MOVIMIENTOS"
                        if summary["documents"]
                        else "SIN MOVIMIENTOS EN ACDOCA"
                    ),
                    account_control,
                    source_row,
                ]
            )

    acdoca_only_keys = sorted(
        set(acdoca_documents_by_asset) - set(ar02_by_key)
    )
    solo_acdoca_headers = [
        "ANLN1",
        "ANLN2",
        "CLAVE",
        "N_DOCUMENTOS",
        "DOCUMENTOS",
        "CUENTAS 32/33 ACDOCA",
        "FILAS ACDOCA",
    ]
    solo_acdoca_records = [
        [
            asset_code,
            subnumber,
            f"{asset_code}-{subnumber}",
            len(acdoca_documents_by_asset[(asset_code, subnumber)]["documents"]),
            " / ".join(
                acdoca_documents_by_asset[(asset_code, subnumber)]["documents"]
            ),
            " / ".join(
                acdoca_documents_by_asset[(asset_code, subnumber)]["accounts"]
            ),
            " / ".join(
                str(source_row)
                for source_row, _ in acdoca_documents_by_asset[
                    (asset_code, subnumber)
                ]["movements"]
            ),
        ]
        for asset_code, subnumber in acdoca_only_keys
    ]
    account_review_count = sum(
        1
        for (asset_code, subnumber), asset in acdoca_documents_by_asset.items()
        if (asset_code, subnumber) in ar02_by_key
        and any(
            account
            != normalize_account(
                ar02_by_key[(asset_code, subnumber)][0][1][
                    ar02_column_indices["cta_cap"]
                ]
            )
            for account in asset["accounts"]
        )
    )
    validation_counts = Counter(
        validation["status"] for validation in validations_by_base_row.values()
    )
    base_without_movements = sum(
        validation["status"] == "VALIDADO AR02"
        and ar02_output_by_base_row[source_row][1]
        == "SIN MOVIMIENTOS EN ACDOCA"
        for source_row, validation in validations_by_base_row.items()
    )
    base_account_review = sum(
        row[AR02_OUTPUT_HEADERS.index("CONTROL CUENTA AR02")]
        == "REVISAR CUENTA"
        for row in ar02_output_by_base_row.values()
    )
    base_disposed_in_ar02 = sum(
        row[AR02_OUTPUT_HEADERS.index("ESTADO ACTIVO AR02")] == "BAJA EN AR02"
        for row in ar02_output_by_base_row.values()
    )
    base_fully_depreciated_in_ar02 = sum(
        validations_by_base_row[source_row]["status"] == "VALIDADO AR02"
        and amount_key(
            row[AR02_OUTPUT_HEADERS.index("CAP ACTUALES AR02")]
        )
        not in (None, Decimal(0))
        and amount_key(
            row[AR02_OUTPUT_HEADERS.index("VALCONT ACTUAL AR02")]
        )
        == 0
        for source_row, row in ar02_output_by_base_row.items()
    )
    base_with_validated_asset = validation_counts["VALIDADO AR02"]
    empty_validation_statuses = [
        source_row
        for source_row, validation in validations_by_base_row.items()
        if not validation["status"].strip()
    ]
    if empty_validation_statuses:
        raise ValueError(
            "Hay filas BASE sin estado de validación AR02: "
            f"{empty_validation_statuses[:10]}."
        )
    acdoca_year_counts = Counter(
        normalize_key(record[acdoca_column_indices["gjahr"]])
        for _, record in acdoca_records
    )
    acdoca_account_prefix_counts = Counter(
        normalize_account(record[acdoca_column_indices["racct"]])[:2]
        for _, record in acdoca_records
    )
    ar02_summary_records = [
        ["Activos únicos en AR02", len(ar02_by_key)],
        [
            "Activos AR02 con movimientos en ACDOCA",
            sum(
                bool(acdoca_documents_by_asset.get(key, {}).get("documents"))
                for key in ar02_by_key
            ),
        ],
        [
            "Activos AR02 sin movimientos en ACDOCA",
            sum(
                not bool(
                    acdoca_documents_by_asset.get(key, {}).get("documents")
                )
                for key in ar02_by_key
            ),
        ],
        ["Activos con cuenta a revisar", account_review_count],
        [
            "Claves duplicadas en AR02",
            sum(len(rows) > 1 for rows in ar02_by_key.values()),
        ],
        ["Activos en ACDOCA que no están en AR02", len(acdoca_only_keys)],
        ["Líneas ACDOCA sin activo", acdoca_without_asset],
        [
            "Filas BASE con activo validado en AR02",
            base_with_validated_asset,
        ],
        [
            "Filas BASE sin activo validado en AR02",
            len(base_records) - base_with_validated_asset,
        ],
        [
            "Filas BASE para revisar subnúmero",
            validation_counts["REVISAR SUBNÚMERO"],
        ],
        [
            "Filas BASE con activo que no existe en AR02",
            validation_counts["NO EXISTE EN AR02"],
        ],
        [
            "Filas BASE con equipo duplicado en IH08",
            validation_counts["REVISAR EQUIPO DUPLICADO"],
        ],
        [
            "Filas BASE sin código de equipo",
            validation_counts["BASE SIN CÓDIGO DE EQUIPO"],
        ],
        [
            "Equipos que no existen en IH08",
            validation_counts["EQUIPO NO EXISTE EN IH08"],
        ],
        [
            "Equipos sin versión vigente en IH08",
            validation_counts["EQUIPO SIN VERSIÓN VIGENTE EN IH08"],
        ],
        [
            "Equipos sin activo fijo en IH08",
            validation_counts["EQUIPO SIN ACTIVO FIJO EN IH08"],
        ],
        ["Filas BASE sin movimientos en ACDOCA", base_without_movements],
        ["Filas BASE con cuenta a revisar", base_account_review],
        ["Filas BASE marcadas BAJA EN AR02", base_disposed_in_ar02],
        [
            "Filas BASE totalmente depreciadas según AR02",
            base_fully_depreciated_in_ar02,
        ],
        ["Filas BASE recuperadas gracias a AR02", len(recovered_base_rows)],
        [
            "Filas BASE con candidato AR02 puntaje >= 90",
            candidate_high_score_count,
        ],
    ]
    def ar02_column(row: list[Any], header: str) -> Any:
        return row[AR02_OUTPUT_HEADERS.index(header)]

    legacy_route_rows = [
        source_row
        for source_row, validation in validations_by_base_row.items()
        if validation.get("detail", "").startswith("Antes:")
    ]
    legacy_route_counts = Counter(
        re.sub(r"^PLACA .* → AR02$", "PLACA → AR02", validations_by_base_row[row]["route"])
        for row in legacy_route_rows
    )
    legacy_comparison_counts = Counter(
        ar02_column(row, "COMPARACIÓN AR02 HANA VS ANTIGUO")
        for row in ar02_output_by_base_row.values()
        if ar02_column(row, "COMPARACIÓN AR02 HANA VS ANTIGUO")
    )
    legacy_state_counts = Counter(
        ar02_column(row, "ESTADO AR02 ANTIGUO").split(" (")[0]
        for row in ar02_output_by_base_row.values()
        if ar02_column(row, "ESTADO AR02 ANTIGUO")
    )
    ar02_summary_records.extend(
        [
            [
                "Columna AF antiguo en AR02 HANA",
                ar02_headers[legacy_column] if legacy_column is not None else "NO ENCONTRADA",
            ],
            ["Activos HANA con AF antiguo (R3)", len(legacy_by_hana)],
            [
                "Columna placa antigua en IH08",
                export_headers[ih08_plate_index]
                if ih08_plate_index is not None
                else "NO ENCONTRADA",
            ],
            [
                "Archivo AR02 ANTIGUO",
                ar02_legacy_path.name if ar02_legacy_path else "NO ENCONTRADO",
            ],
            ["Activos (R3 + subnúmero) en AR02 ANTIGUO", len(ar02_legacy)],
            [
                "Filas BASE recuperadas por rutas nuevas (AF R3 / placa)",
                len(legacy_route_rows),
            ],
            *[
                [f"  Ruta {route}", count]
                for route, count in sorted(legacy_route_counts.items())
            ],
            ["Placas en AR02 HANA", len(ar02_by_plate)],
            *[
                [f"Filas BASE - AR02 ANTIGUO {state}", count]
                for state, count in sorted(legacy_state_counts.items())
            ],
            *[
                [f"Filas BASE - comparación HANA vs antiguo: {state}", count]
                for state, count in sorted(legacy_comparison_counts.items())
            ],
        ]
    )
    ar02_summary_records.extend(
        [
            [
                f"Filas ACDOCA incluidas - ejercicio {year or '(vacío)'}",
                count,
            ]
            for year, count in sorted(acdoca_year_counts.items())
        ]
    )
    ar02_summary_records.extend(
        [
            [
                f"Filas ACDOCA incluidas - prefijo cuenta {prefix or '(vacío)'}",
                count,
            ]
            for prefix, count in sorted(acdoca_account_prefix_counts.items())
        ]
    )
    if not OUTPUT_FOLDER.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de salida: {OUTPUT_FOLDER}")

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    base_end = len(visible_base_headers)
    export_start = base_end + 1
    export_end = base_end + len(visible_export_headers)
    cross_status_column = base_end + len(visible_export_headers) + 1
    acdoca_start_column = cross_status_column + 1
    acdoca_end_column = acdoca_start_column + len(ACDOCA_OUTPUT_HEADERS) - 1
    ar02_start_column = acdoca_end_column + 1
    ar02_end_column = len(merged_headers)
    status_headers = STATUS_HEADERS + ACDOCA_OUTPUT_HEADERS + AR02_OUTPUT_HEADERS
    acdoca_status_counts = Counter(
        row[merged_headers.index("ESTADO CONCILIACIÓN ACDOCA")]
        for row in merged_records
    )
    completed_reconciliations = [
        reconciliation
        for reconciliation in reconciliations_by_base_row.values()
        if all(value is not None for value in reconciliation[0])
    ]
    reconciliation_route_counts = Counter(
        reconciliation[4] for reconciliation in completed_reconciliations
    )
    add_records_sheet(
        workbook,
        "STATUS BAJA 2026",
        status_headers,
        status_records,
        "StatusBaja2026",
        (
            f"Formato de campos basado en el archivo de referencia (solo lectura)  |  "
            f"Base: {base_path.name}  |  {len(status_records)} registros  |  "
            f"ACDOCA única: {acdoca_status_counts['Coincidencia única']}  |  "
            f"BKPF/ACDOCA: {reconciliation_route_counts['BKPF/MKPF → ACDOCA']}  |  "
            f"múltiples: {sum(count for status, count in acdoca_status_counts.items() if status.startswith('Múltiples movimientos'))}"
        ),
        [
            (1, 8, "CONTROL Y REFERENCIA", BLUE),
            (9, 21, "IDENTIFICACIÓN DEL ACTIVO", TEAL),
            (22, 25, "BAJA Y CONTABILIDAD", BLUE),
            (26, len(STATUS_HEADERS), "SEGUIMIENTO", GOLD),
            (
                len(STATUS_HEADERS) + 1,
                len(STATUS_HEADERS) + len(ACDOCA_OUTPUT_HEADERS),
                "CONCILIACIÓN ACDOCA/BKPF",
                TEAL,
            ),
            (
                len(STATUS_HEADERS) + len(ACDOCA_OUTPUT_HEADERS) + 1,
                len(status_headers),
                "VALIDACIÓN AR02",
                GOLD,
            ),
        ],
        create_table=False,
    )
    add_records_sheet(
        workbook,
        "CANDIDATOS AR02",
        candidate_headers,
        candidate_records,
        "CandidatosAR02",
        (
            f"{len(candidate_top_scores)} filas BASE en revisión; "
            f"{candidate_high_score_count} con puntaje >= 90. "
            "Solo sugerencias por similitud, no asignadas al resultado."
        ),
        [(1, len(candidate_headers), "SUGERENCIAS NO ASIGNADAS", GOLD)],
    )
    add_records_sheet(
        workbook,
        "RESUMEN AR02",
        ["Indicador", "Cantidad"],
        ar02_summary_records,
        "ResumenAR02",
        (
            f"AR02: {ar02_path.name}  |  ACDOCA: {acdoca_path.name}  |  "
            f"BASE: {len(base_records)} filas · "
            f"{len(recovered_base_rows)} recuperadas por AR02"
        ),
        [(1, 2, "CONTEOS DEL CRUCE", BLUE)],
    )
    add_records_sheet(
        workbook,
        "AR02 completo",
        ar02_sheet_headers,
        ar02_sheet_records,
        "ActivosAR02",
        (
            f"Lista maestra de activos: {ar02_path.name}  |  "
            f"{len(ar02_sheet_records)} filas válidas · "
            f"{len(ar02_by_key)} claves ANLN1/ANLN2"
        ),
        [
            (1, len(ar02_headers), "DATOS ORIGINALES AR02", BLUE),
            (
                len(ar02_headers) + 1,
                len(ar02_sheet_headers),
                "MOVIMIENTOS ACDOCA Y CONTROLES",
                TEAL,
            ),
        ],
    )
    add_records_sheet(
        workbook,
        "SOLO EN ACDOCA",
        solo_acdoca_headers,
        solo_acdoca_records,
        "ActivosSoloACDOCA",
        f"Activos con movimientos ACDOCA que no aparecen en AR02: {len(solo_acdoca_records)}",
        [(1, len(solo_acdoca_headers), "REVISAR ACTIVOS FUERA DE AR02", GOLD)],
    )
    add_records_sheet(
        workbook,
        "ACDOCA - Revisión múltiple",
        [
            "ID BASE",
            "FILA BASE ORIGEN",
            "CODIGO AF HANAH",
            "FILA ACDOCA ORIGEN",
        ]
        + [f"ACDOCA - {header}" for header in acdoca_headers],
        acdoca_review_records,
        "ACDOCARevisionMultiple",
        (
            f"{len(acdoca_review_records)} movimientos asociados a activos con "
            "más de una coincidencia. No se escogió ni sumó ningún movimiento. "
            "La conciliación de activos usa ANLN1 y ANLN2 después de validar "
            "el subnúmero contra AR02."
        ),
        [
            (1, 4, "REFERENCIA AL REGISTRO BASE", GOLD),
            (5, 13, "DOCUMENTO Y CONTABILIDAD", TEAL),
            (14, len(acdoca_headers) + 4, "REGISTRO ACDOCA ORIGINAL", BLUE),
        ],
    )
    add_records_sheet(
        workbook,
        "Bajas completas",
        merged_headers,
        merged_records,
        "BajasCompletas",
        (
            f"Base: {base_path.name}  |  IH08: {export_path.name}  |  "
            f"BKPF: {', '.join(path.name for path in bkpf_paths)}  |  "
            f"{len(base_records)} registros · {matched_count} coincidencias · "
            f"{len(unmatched_records)} pendientes de revisar"
        ),
        [
            (1, base_end, "INFORMACIÓN DE LA BAJA", BLUE),
            (
                export_start,
                export_end,
                "DATOS DEL EXPORT IH08",
                TEAL,
            ),
            (cross_status_column, cross_status_column, "VALIDACIÓN", GOLD),
            (
                acdoca_start_column,
                acdoca_end_column,
                "CONCILIACIÓN SAP ACDOCA/BKPF",
                TEAL,
            ),
            (
                ar02_start_column,
                ar02_end_column,
                "VALIDACIÓN DEL ACTIVO AR02",
                GOLD,
            ),
        ],
    )
    legacy_sheet_headers = [
        "ID BASE",
        "FILA BASE",
        "EQUIPO",
        "ESTADO VALIDACIÓN AR02",
        "CLAVE ACTIVO AR02",
        "DENOMINACIÓN AR02",
        "CAP ACTUALES AR02",
        "VALCONT ACTUAL AR02",
        "DETALLE VALIDACIÓN",
        *AR02_OUTPUT_HEADERS[AR02_OUTPUT_HEADERS.index("RUTA VALIDACIÓN AR02"):],
    ]
    legacy_sheet_records: list[list[Any]] = []
    for base_record, base_source_row in zip(base_records, base_source_rows):
        row = ar02_output_by_base_row[base_source_row]
        validation = validations_by_base_row[base_source_row]
        if not (
            ar02_column(row, "AF R3 (ANTIGUO)")
            or ar02_column(row, "PLACA ANTIGUA IH08")
            or validation.get("detail", "").startswith("Antes:")
        ):
            continue
        legacy_sheet_records.append(
            [
                source_value(base_record, base_headers, "ID"),
                base_source_row,
                cell_text(base_record[BASE_EQUIPMENT_COLUMN - 1])
                if len(base_record) >= BASE_EQUIPMENT_COLUMN
                else "",
                ar02_column(row, "ESTADO VALIDACIÓN AR02"),
                ar02_column(row, "CLAVE ACTIVO AR02"),
                ar02_column(row, "DENOMINACIÓN AR02"),
                ar02_column(row, "CAP ACTUALES AR02"),
                ar02_column(row, "VALCONT ACTUAL AR02"),
                validation.get("detail", ""),
                *row[AR02_OUTPUT_HEADERS.index("RUTA VALIDACIÓN AR02"):],
            ]
        )
    add_records_sheet(
        workbook,
        "AR02 ANTIGUO - CRUCE",
        legacy_sheet_headers,
        legacy_sheet_records,
        "CruceAR02Antiguo",
        (
            f"AR02 ANTIGUO: "
            f"{ar02_legacy_path.name if ar02_legacy_path else 'no encontrado'}  |  "
            f"{len(legacy_sheet_records)} filas BASE con placa antigua / AF R3  |  "
            f"{len(legacy_route_rows)} recuperadas por rutas nuevas"
        ),
        [
            (1, 9, "BAJA / ACTIVO HANA", BLUE),
            (10, len(legacy_sheet_headers), "PLACA ANTIGUA / AR02 ANTIGUO (R3)", GOLD),
        ],
    )
    add_records_sheet(
        workbook,
        "IH08 completo",
        visible_export_headers,
        [
            [record[index] for index in export_column_indices]
            for record in export_records
        ],
        "IH08Completo",
        f"Exportación original: {export_path.name}  |  {len(export_records)} registros",
        [(1, len(visible_export_headers), "DATOS DEL EXPORT IH08", TEAL)],
    )
    add_records_sheet(
        workbook,
        "Bajas sin IH08",
        visible_base_headers
        + ["Resultado del cruce"]
        + ACDOCA_OUTPUT_HEADERS
        + AR02_OUTPUT_HEADERS,
        unmatched_records,
        "BajasSinIH08",
        f"Revisión manual requerida  |  {len(unmatched_records)} registros sin coincidencia",
        [
            (1, base_end, "INFORMACIÓN DE LA BAJA", BLUE),
            (base_end + 1, base_end + 1, "VALIDACIÓN", GOLD),
            (
                base_end + 2,
                base_end + 1 + len(ACDOCA_OUTPUT_HEADERS),
                "CONCILIACIÓN SAP ACDOCA",
                TEAL,
            ),
            (
                base_end + 2 + len(ACDOCA_OUTPUT_HEADERS),
                base_end
                + 1
                + len(ACDOCA_OUTPUT_HEADERS)
                + len(AR02_OUTPUT_HEADERS),
                "VALIDACIÓN AR02",
                GOLD,
            ),
        ],
    )
    add_records_sheet(
        workbook,
        "BKPF - Auditoría",
        audit_headers,
        bkpf_audit_records,
        "ConciliacionBKPFAuditoria",
        (
            f"BKPF: {', '.join(path.name for path in bkpf_paths)}  |  "
            f"{len(bkpf_audit_records)} filas de evidencia  |  "
            "AWKEY se valida como MBLNR (10) + MJAHR (4); "
            "no se usa XBLNR como clave de conciliación."
        ),
        [
            (1, 8, "REGISTRO BASE Y DECISIÓN", GOLD),
            (9, 12, "DOCUMENTO MATERIAL / AWKEY", TEAL),
            (13, 20, "CABECERA BKPF", BLUE),
            (21, len(audit_headers), "POSICIONES ACDOCA", TEAL),
        ],
    )
    unmatched_sheet = workbook["Bajas completas"]
    status_letter = openpyxl.utils.get_column_letter(cross_status_column)
    unmatched_sheet.conditional_formatting.add(
        f"{status_letter}5:{status_letter}{len(merged_records) + 4}",
        FormulaRule(
            formula=[f'{status_letter}5="Sin coincidencia"'],
            fill=PatternFill("solid", fgColor="FCE4D6"),
        ),
    )
    unmatched_sheet.conditional_formatting.add(
        f"{status_letter}5:{status_letter}{len(merged_records) + 4}",
        FormulaRule(
            formula=[f'{status_letter}5="Coincide"'],
            fill=PatternFill("solid", fgColor="E2F0D9"),
        ),
    )
    workbook["IH08 completo"].sheet_properties.tabColor = TEAL
    workbook["Bajas sin IH08"].sheet_properties.tabColor = GOLD
    workbook["ACDOCA - Revisión múltiple"].sheet_properties.tabColor = GOLD
    workbook["BKPF - Auditoría"].sheet_properties.tabColor = BLUE
    accounting_status_letter = openpyxl.utils.get_column_letter(
        merged_headers.index("ESTADO CONCILIACIÓN ACDOCA") + 1
    )
    unmatched_sheet.conditional_formatting.add(
        f"{accounting_status_letter}5:{accounting_status_letter}{len(merged_records) + 4}",
        FormulaRule(
            formula=[f'{accounting_status_letter}5="Coincidencia única"'],
            fill=PatternFill("solid", fgColor="E2F0D9"),
        ),
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    status_result = update_status_template(
        status_candidates,
        STATUS_TEMPLATE_PATH,
        STATUS_OUTPUT_FOLDER,
        timestamp,
        resolve_equipment,
        ih08_fields,
        asset_book_value,
        asset_denomination,
        baja_documents_by_asset,
        hana_codes_from_r3,
        legacy_book_info,
        r3_legacy_detail,
    )
    add_records_sheet(
        workbook,
        "STATUS - TRAZABILIDAD",
        STATUS_TRACE_HEADERS,
        status_result["trace"],
        "StatusTrazabilidad",
        (
            f"Cada celda nueva escrita en {status_result['output_path'].name} "
            "(en amarillo) y de dónde salió."
        ),
        [(1, len(STATUS_TRACE_HEADERS), "CELDAS COMPLETADAS EN EL STATUS", GOLD)],
    )
    add_records_sheet(
        workbook,
        "STATUS - REVISAR",
        STATUS_REVIEW_HEADERS,
        status_result["review"],
        "StatusRevisar",
        (
            "Casos que no se llenaron o que contradicen SAP. "
            "Nada de esto se modificó en el STATUS."
        ),
        [(1, len(STATUS_REVIEW_HEADERS), "REVISIÓN MANUAL", GOLD)],
    )
    workbook["STATUS - TRAZABILIDAD"].sheet_properties.tabColor = "FFFF00"
    workbook["STATUS - REVISAR"].sheet_properties.tabColor = GOLD
    workbook.move_sheet("STATUS - REVISAR", offset=-(len(workbook.sheetnames) - 2))
    workbook.move_sheet(
        "STATUS - TRAZABILIDAD", offset=-(len(workbook.sheetnames) - 2)
    )

    output_path = OUTPUT_FOLDER / f"Resultado_completo_{timestamp}.xlsx"
    suffix = 1
    while output_path.exists():
        output_path = OUTPUT_FOLDER / f"Resultado_completo_{timestamp}_{suffix:02d}.xlsx"
        suffix += 1
    workbook.save(output_path)

    print(f"Base utilizada: {base_path.name}")
    print(f"Export utilizado: {export_path.name}")
    print(f"AR02 utilizado: {ar02_path.name}")
    print(
        "Columna AF antiguo en AR02: "
        + (ar02_headers[legacy_column] if legacy_column is not None else "NO ENCONTRADA")
        + f" ({len(legacy_by_hana)} activos HANA con código R3)"
    )
    print(
        "Columna placa antigua en IH08: "
        + (
            export_headers[ih08_plate_index]
            if ih08_plate_index is not None
            else "NO ENCONTRADA"
        )
    )
    print(
        "AR02 ANTIGUO utilizado: "
        + (ar02_legacy_path.name if ar02_legacy_path else "NO ENCONTRADO")
        + f" ({len(ar02_legacy)} activos)"
    )
    print(
        "Filas BASE recuperadas por rutas nuevas (AF R3 / placa): "
        f"{len(legacy_route_rows)}"
    )
    for route, count in sorted(legacy_route_counts.items()):
        print(f"  {route}: {count}")
    for state, count in sorted(legacy_comparison_counts.items()):
        print(f"  Comparación HANA vs antiguo - {state}: {count}")
    print(f"ACDOCA utilizado: {acdoca_path.name}")
    print(f"BKPF utilizado: {', '.join(path.name for path in bkpf_paths)}")
    print(f"Resultado creado: {output_path}")
    print(f"Status creado: {status_result['output_path']}")
    print(
        f"Registros de BASE ya presentes en el STATUS: "
        f"{status_result['skipped_existing']}"
    )
    print(f"Filas de BASE agregadas al STATUS: {status_result['appended_rows']}")
    print(
        "Filas existentes del STATUS completadas: "
        f"{status_result['filled_existing_rows']}"
    )
    print("Celdas STATUS completadas por columna:")
    for column, count in sorted(status_result["fill_counts"].items()):
        print(f"  {column}: {count}")
    print(f"Total de celdas STATUS completadas: {status_result['filled_cells']}")
    print(
        "Celdas STATUS preexistentes modificadas: 0 "
        f"(verificado celda por celda hasta la fila {status_result['verified_cells']})"
    )
    review_counts = Counter(row[3] for row in status_result["review"])
    print("Casos para revisar en la hoja STATUS - REVISAR:")
    for alert, count in review_counts.most_common():
        print(f"  {alert}: {count}")
    print("Ejemplos de celdas STATUS completadas:")
    for row_number, column, value in status_result["examples"]:
        print(f"  fila {row_number}, {column}: {value}")
    print(f"Registros de la base: {len(base_records)}")
    print(f"Coincidencias por código de equipo: {matched_count}")
    print(f"Pendientes de revisar: {len(unmatched_records)}")
    print(
        "Estado contable según AR02: "
        f"{base_disposed_in_ar02} BAJA EN AR02; "
        f"{base_fully_depreciated_in_ar02} totalmente depreciadas."
    )
    print(
        "Registros ACDOCA después de filtros AG10 / 0L / 32* / 33* "
        "en todos los ejercicios: "
        f"{len(acdoca_records)}"
    )
    print("Filas ACDOCA incluidas por ejercicio:")
    for year, count in sorted(acdoca_year_counts.items()):
        print(f"  {year or '(vacío)'}: {count}")
    print("Filas ACDOCA incluidas por prefijo de cuenta:")
    for prefix, count in sorted(acdoca_account_prefix_counts.items()):
        print(f"  {prefix or '(vacío)'}*: {count}")
    print(
        "Conciliación ACDOCA: "
        f"{acdoca_status_counts['Coincidencia única']} únicas, "
        f"{acdoca_status_counts['Sin coincidencia ACDOCA']} sin coincidencia, "
        f"{acdoca_status_counts['Pendiente: sin código HANA']} pendientes sin HANA, "
        f"{sum(count for status, count in acdoca_status_counts.items() if status.startswith('Múltiples movimientos'))} activos con múltiples movimientos, "
        f"{len(acdoca_review_records)} movimientos en hoja auxiliar."
    )
    print(
        "Conciliaciones contables completas: "
        f"{len(completed_reconciliations)} total "
        f"({reconciliation_route_counts['ANLN1 + ANLN2']} por ANLN1+ANLN2, "
        f"{reconciliation_route_counts['BKPF/MKPF → ACDOCA']} por BKPF/MKPF)."
    )
    print(f"Filas de evidencia en BKPF - Auditoría: {len(bkpf_audit_records)}")
    print(
        "Validación BASE/AR02: "
        f"{validation_counts['VALIDADO AR02']} validadas, "
        f"{len(base_records) - validation_counts['VALIDADO AR02']} sin activo "
        f"validado; {len(recovered_base_rows)} recuperadas."
    )
    print(
        "Marcas BASE: "
        f"NO EXISTE EN AR02={validation_counts['NO EXISTE EN AR02']}, "
        f"REVISAR SUBNÚMERO={validation_counts['REVISAR SUBNÚMERO']}, "
        f"REVISAR EQUIPO DUPLICADO="
        f"{validation_counts['REVISAR EQUIPO DUPLICADO']}, "
        f"SIN MOVIMIENTOS EN ACDOCA={base_without_movements}, "
        f"REVISAR CUENTA={base_account_review}."
    )
    print("Ejemplos de filas BASE recuperadas (máximo 10):")
    for example in recovered_base_rows[:10]:
        print(
            f"  ID={example[0]!r}, fila={example[1]}, Equipo={example[2]!r}, "
            f"ANLN1={example[3]}, ANLN2={example[4]}, "
            f"denominación={example[5]!r}"
        )
    print(
        "Candidatos AR02: "
        f"{len(candidate_top_scores)} filas revisadas, "
        f"{candidate_high_score_count} con puntaje >= 90."
    )
    print("Ejemplos de candidatos AR02 (máximo 10 filas BASE):")
    for candidate_record in candidate_examples[:10]:
        print(
            f"  ID={candidate_record[0]!r}, descripción="
            f"{candidate_record[1]!r}, activo={candidate_record[2]}, "
            f"denominación={candidate_record[3]!r}, "
            f"puntaje={candidate_record[4]}, centro={candidate_record[5]!r}"
        )
    print(f"Filas con candidato >= 90: {candidate_high_score_count}")
    return output_path


if __name__ == "__main__":
    try:
        generate()
    except Exception as error:
        print(f"ERROR: {error}")
        raise SystemExit(1) from error
