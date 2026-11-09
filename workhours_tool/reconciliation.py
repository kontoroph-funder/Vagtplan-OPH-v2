"""Snapshot reconciliation of exported person worksheets; no shift adjustments."""
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO
import posixpath
import re
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.datetime import from_excel

HEADERS = [
    "Dato", "Dag", "Starttid", "Sluttid", "Vagtform", "Institution",
    "Personsnavn", "Vikar/Fastansat", "Summering efter reduktion",
    "Fulde vagtforms navn", "Bemærkning", "Afvigelse",
]
SUMMARY_NAMES = {"samlet", "samlet oversigts"}
DAYS = ("mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag", "søndag")
SIGNED = '+0.00;-0.00;0.00'


class ReconciliationError(ValueError):
    """An actionable workbook validation failure."""


def _normalize(value):
    return " ".join(str(value or "").split()).casefold()


def _header(value):
    value = _normalize(value)
    if value == "summering efter reduktioner":
        return "summering efter reduktion"
    if value.replace(" ", "") in {"vikar", "vikar/fastansat"}:
        return "vikar/fastansat"
    return value


def _date(value, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.date()
    if isinstance(value, str):
        for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    raise ValueError("ugyldig dato")


def _time(value, epoch=None):
    if value is None or value == "":
        return 86401, None
    if isinstance(value, datetime):
        # A numeric Excel time of 1 formatted as hh:mm loads as epoch day 1.
        if epoch is not None and value == from_excel(1, epoch):
            return 86400, "24:00"
        value = value.time()
    if isinstance(value, time):
        seconds = value.hour * 3600 + value.minute * 60 + value.second
    elif isinstance(value, timedelta):
        seconds = value.total_seconds()
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        seconds = round(value * 86400, 6)
    elif isinstance(value, str) and re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", value.strip()):
        parts = [int(p) for p in value.strip().split(":")]
        h, m = parts[:2]
        s = parts[2] if len(parts) == 3 else 0
        if m >= 60 or s >= 60:
            raise ValueError("ugyldigt klokkeslæt")
        seconds = h * 3600 + m * 60 + s
    else:
        raise ValueError("ugyldigt klokkeslæt")
    if not 0 <= seconds <= 86400:
        raise ValueError("ugyldigt klokkeslæt")
    seconds = int(seconds)
    h, remainder = divmod(seconds, 3600)
    m, s = divmod(remainder, 60)
    return seconds, f"{h:02}:{m:02}" + (f":{s:02}" if s else "")


def _hours(value):
    if isinstance(value, bool) or value is None:
        raise ValueError("ulæseligt timetal")
    try:
        result = Decimal(str(value).strip().replace(",", "."))
    except InvalidOperation as exc:
        raise ValueError("ulæseligt timetal") from exc
    if not result.is_finite():
        raise ValueError("ulæseligt timetal")
    return result


def _preserve_formula_caches(original, generated):
    """openpyxl drops cached results on save; retain them for repeat uploads."""
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def sheets(archive):
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: posixpath.normpath(posixpath.join("xl", r.attrib["Target"]))
                   if not r.attrib["Target"].startswith("/") else r.attrib["Target"].lstrip("/")
                   for r in relationships}
        root = ET.fromstring(archive.read("xl/workbook.xml"))
        return {s.attrib["name"]: targets[s.attrib[f"{{{rel_ns}}}id"]]
                for s in root.findall("s:sheets/s:sheet", ns)}

    output = BytesIO()
    with ZipFile(BytesIO(original)) as source, ZipFile(BytesIO(generated)) as new:
        source_sheets, new_sheets = sheets(source), sheets(new)
        replacements = {}
        for name, path in source_sheets.items():
            if name.casefold() in SUMMARY_NAMES or name not in new_sheets:
                continue
            source_root = ET.fromstring(source.read(path))
            caches = {}
            for cell in source_root.findall(".//s:sheetData/s:row/s:c", ns):
                value = cell.find("s:v", ns)
                if cell.find("s:f", ns) is not None and value is not None and value.text is not None:
                    caches[cell.attrib["r"]] = (value, cell.attrib.get("t"))
            if not caches:
                continue
            target_path = new_sheets[name]
            root = ET.fromstring(new.read(target_path))
            for cell in root.findall(".//s:sheetData/s:row/s:c", ns):
                if cell.attrib["r"] not in caches or cell.find("s:f", ns) is None:
                    continue
                cached, data_type = caches[cell.attrib["r"]]
                previous = cell.find("s:v", ns)
                if previous is not None:
                    cell.remove(previous)
                cell.append(cached)
                if data_type:
                    cell.set("t", data_type)
                else:
                    cell.attrib.pop("t", None)
            replacements[target_path] = ET.tostring(root)
        with ZipFile(output, "w", ZIP_DEFLATED) as result:
            for entry in new.infolist():
                result.writestr(entry, replacements.get(entry.filename, new.read(entry.filename)))
    return output.getvalue()


def _read_rows(workbook, cached):
    result = []
    required = {_header(h) for h in HEADERS[:11]} - {"dag"}
    for sheet in workbook:
        if sheet.title.casefold() in SUMMARY_NAMES:
            continue
        columns = {}
        header_row = None
        for cells in sheet.iter_rows():
            names = [_header(c.value) for c in cells]
            if "dato" not in names:
                continue
            for cell, name in zip(cells, names):
                if name in required | {"dag"}:
                    if name in columns:
                        raise ReconciliationError(f"{sheet.title}!{cell.coordinate}: tvetydig kolonne {cell.value}.")
                    columns[name] = cell.column
            header_row = cells[0].row
            break
        missing = required - columns.keys()
        if missing:
            raise ReconciliationError(f"{sheet.title}: manglende nødvendige kolonner: {', '.join(sorted(missing))}.")
        for cells in sheet.iter_rows(min_row=header_row + 1):
            if all(c.value is None or c.value == "" for c in cells):
                continue
            date_cell = cells[columns["dato"] - 1]
            if _normalize(date_cell.value) in {"total", "i alt", "totaler", "ingen vagter fundet i den angivne html."}:
                continue

            def read(name, parser=None):
                cell = cells[columns[name] - 1]
                value = cell.value
                if cell.data_type == "f":
                    value = cached[sheet.title][cell.coordinate].value
                    if value is None:
                        raise ReconciliationError(
                            f"{sheet.title}!{cell.coordinate}: formlen mangler et gemt resultat. "
                            "Genberegn og gem filen i Excel, og upload den igen."
                        )
                try:
                    return parser(value) if parser else value
                except (ValueError, TypeError, OverflowError) as exc:
                    raise ReconciliationError(f"{sheet.title}!{cell.coordinate}: {exc} ({value!r}).") from exc

            day = read("dato", lambda v: _date(v, workbook.epoch))
            start_key, start = read("starttid", lambda v: _time(v, workbook.epoch))
            end_key, end = read("sluttid", lambda v: _time(v, workbook.epoch))
            marker = read("vikar/fastansat")
            hours = read("summering efter reduktion", _hours)
            values = [day, DAYS[day.weekday()], start, end, read("vagtform"),
                      read("institution"), read("personsnavn"),
                      "Fastansat" if _normalize(marker) in {"", "fastansat"} else "Vikar",
                      hours, read("fulde vagtforms navn"), read("bemærkning")]
            result.append(((day, start_key, end_key, str(values[6] or "").casefold()), values))
    return [values for _, values in sorted(result, key=lambda item: item[0])]


def reconcile_workbook(workbook_bytes, mode) -> bytes:
    """Return a new XLSX containing preserved person sheets and reconciliation."""
    if mode not in {"1:1", "2:1"}:
        raise ReconciliationError("Ugyldig model: vælg 1:1 eller 2:1.")
    try:
        workbook = load_workbook(BytesIO(workbook_bytes), rich_text=True)
        cached = load_workbook(BytesIO(workbook_bytes), data_only=True)
    except Exception as exc:
        raise ReconciliationError(f"Kunne ikke læse Excel-filen: {exc}") from exc
    rows = _read_rows(workbook, cached)
    if not rows:
        raise ReconciliationError("Ingen vagter fundet")
    for name in list(workbook.sheetnames):
        if name.casefold() in SUMMARY_NAMES:
            del workbook[name]
    sheet = workbook.create_sheet("Samlet")
    sheet.append(HEADERS)
    sheet.freeze_panes = "A2"
    grouped = defaultdict(list)
    for row in rows:
        grouped[row[0]].append(row)
    first, last = rows[0][0], rows[-1][0]
    count = (last - first).days + 1
    norm = Decimal(22 if mode == "1:1" else 44)
    total = Decimal(0)
    thick = Border(top=Side(style="thick", color="000000"))
    for offset in range(count):
        day = first + timedelta(days=offset)
        daily_rows = grouped[day] or [[day, DAYS[day.weekday()], None, None, None,
                                      None, None, None, Decimal(0), None, "Ingen vagter"]]
        daily_total = sum((r[8] for r in daily_rows), Decimal(0))
        total += daily_total
        for index, row in enumerate(daily_rows):
            sheet.append(row + [daily_total - norm if index == len(daily_rows) - 1 else None])
            if offset and index == 0:
                for cell in sheet[sheet.max_row]:
                    cell.border = thick
    sheet.auto_filter.ref = f"A1:L{sheet.max_row}"
    difference = total - norm * count
    sheet.append(["Total"] + [None] * 7 + [total, None, None, difference])
    for cell in sheet[1] + sheet[sheet.max_row]:
        cell.font = Font(bold=True)
    for row in sheet.iter_rows(min_row=2):
        row[0].number_format = "dd.mm.yyyy"
        row[8].number_format = "0.00"
        row[11].number_format = SIGNED
        for cell in row:
            # Source text is a literal snapshot, even when it begins with '='.
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for column, width in zip("ABCDEFGHIJKL", (14, 13, 12, 12, 16, 26, 28, 19, 25, 30, 45, 16)):
        sheet.column_dimensions[column].width = width
    overview = workbook.create_sheet("Samlet Oversigts")
    for label, value in [
        ("Valgt model", mode), ("Startdato", first), ("Slutdato", last),
        ("Antal kalenderdage", count), ("Optimale timer pr. dag", norm),
        ("Optimalt timeforbrug", norm * count), ("Aktuelt timeforbrug", total),
        ("Difference", difference),
    ]:
        overview.append([label, value])
        overview.cell(overview.max_row, 1).font = Font(bold=True)
    for row in (2, 3):
        overview.cell(row, 2).number_format = "dd.mm.yyyy"
    for row in range(5, 9):
        overview.cell(row, 2).number_format = SIGNED if row == 8 else "0.00"
    status, color = (("På niveau", "C6EFCE") if difference == 0 else
                     ("Over optimalt", "FFC7CE") if difference > 0 else
                     ("Under optimalt", "FFEB9C"))
    overview["C8"] = status
    for cell in (overview["B8"], overview["C8"]):
        cell.fill = PatternFill("solid", fgColor=color)
    for column, width in (("A", 28), ("B", 22), ("C", 22)):
        overview.column_dimensions[column].width = width
    output = BytesIO()
    workbook.save(output)
    return _preserve_formula_caches(workbook_bytes, output.getvalue())
