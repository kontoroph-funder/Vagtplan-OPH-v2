from __future__ import annotations

import html
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from zipfile import BadZipFile


KNOWN_WORK_TYPES = {"normtid", "ferie", "bagvagt", "batvagt", "+timer", "-timer", "- timer"}
WORK_TYPE_ALIASES = {
    "vikar omdamsgård": "Normtid",
    "fri med telefon": "Normtid",
    "vikar stn": "Normtid",
    "vikar odamsgård": "Normtid",
}
IGNORED_WORK_TYPES = {
    "-timer",
    "- timer",
    "afspadsering",
    "sove",
    "fri - byttet vagt med kollega",
    "fri - tager en anden dag",
    "ferie uden løn/projektansat",
    "ledelsebagvagt - mike kaldes kontaktes ved behov",
}
WORK_TYPE_ABBREVIATIONS = {
    "+timer": "P",
    "+ timer": "P",
    "bagvagt": "B",
    "batvagt": "N",
    "barn 1. sygedag": "S",
    "barsel": "BA",
    "ferie": "F",
    "helligdag": "N",
    "koloni": "K",
    "kursus": "K",
    "kolonitur": "K",
    "akademi uddannelse": "K",
    "førstehjælp kursus": "K",
    "ledelsesseminar": "N",
    "lederseminar": "N",
    "ledermøde": "N",
    "normtid": "N",
    "p-dag": "N",
    "p-møde": "N",
    "projekt overgreb": "O",
    "projekt a": "N",
    "projekt b": "N",
    "projekt c": "N",
    "syg": "S",
    "administrativ tid": "N",
    "administrativ tid hjemme": "N",
    "møde": "N",
    "mus-samtale": "N",
    "udekørende besøg hos lukas": "N",
    "udekørende besøg hos sigrid": "N",
    "seniordag": "N",
    "marianne": "N",
    "vikar": "N",
    "vikar stegdalsvej": "N",
    "vikar borum": "N",
    "vikar frøstrupskovvej": "N",
    "vikar fsv2": "N",
    "alpha vikar fv2": "N",
    "vikar frøstrupvej1": "N",
    "vikar havndal": "N",
    "vikar lyngbygade": "N",
    "vikar vrangstrupvej": "N",
    "vikar ørnereden": "N",
    "vikar - oph": "N",
    "vikar - aw": "N",
    "vikar feriekolonien": "N",
    "vikar solo": "N",
    "vikar stn": "N",
    "vikar stu": "N",
    "vikar paragraf12 stk3": "N",
    "vikar hesselagergård": "N",
}
DANISH_WEEKDAYS = {
    "Monday": "mandag",
    "Tuesday": "tirsdag",
    "Wednesday": "onsdag",
    "Thursday": "torsdag",
    "Friday": "fredag",
    "Saturday": "lørdag",
    "Sunday": "søndag",
}

INSTITUTION_NAME_OVERRIDES: Dict[str, str] = {
    "Alpha.dk - Borum": "Alpha Borum",
    "Alpha.dk - Frøstrupskovvej1": "Alpha Frøstrupskovvej 1",
    "Alpha.dk - Frøstrupskovvej2": "Alpha Frøstrupskovvej 2",
    "Alpha.dk - Frøstrupvej 2": "Alpha Frøstrupvej 2",
    "Alpha.dk - Frøstrupvej1": "Alpha Frøstrupvej 1",
    "Alpha.dk - Havndal": "Alpha Havndal",
    "Alpha.dk - Lyngbygade": "Alpha Lyngbygade",
    "Alpha.dk - Vrangstrupvej": "Alpha Vrangstrupvej",
    "Alpha.dk - Ørnereden": "Ørnereden",
    "Opholdssteder - Feriekolonien Stenbjerg": "Feriekolonien",
    "Opholdssteder - Hag": "Hesselagergård",
    "Opholdssteder - Odamsgård": "Odamsgård",
    "Opholdssteder - Solo": "Solo",
    "Opholdssteder - Station Nord": "Station Nord",
    "Opholdssteder nr 1 Paragraf 12 stk 3": "Paragraf 12 stk. 3",
    "Opholdssteder nr 2 Paragraf 12 stk 3": "Paragraf 12 stk. 3",
    "Opholdssteder nr 3 Paragraf 12 stk 3": "Paragraf 12 stk. 3",
    "Opholdssteder nr 4 Paragraf 12 stk 3": "Paragraf 12 stk. 3",
    "Opholdssteder nr 5 Paragraf 12 stk 3": "Paragraf 12 stk. 3",
}


def parse_shift_rows(html: str, person_name: str | None = None, institution_overrides: Dict[str, str] | None = None) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    rows: List[Dict[str, Any]] = []

    for element in soup.select(".dbar"):
        data_t0 = element.get("data-t0") or ""
        data_t1 = element.get("data-t1") or ""
        faux_title = element.get("faux-title") or ""
        data_person = element.get("data-person") or person_name or ""

        if not data_t0 or not data_t1:
            continue

        start_dt = datetime.strptime(data_t0, "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(data_t1, "%Y-%m-%d %H:%M")

        if end_dt < start_dt:
            end_dt = end_dt + timedelta(days=1)

        work_type = extract_work_type(element, faux_title) or "Normtid"
        normalized_work_type_input = work_type.strip().lower()
        if normalized_work_type_input in {"fri med telefon", "-timer", "- timer"}:
            continue
        normalized_work_type = normalize_work_type(work_type)
        if not normalized_work_type:
            continue

        comments = ""
        nordic_pas_minus_hours = None
        if faux_title:
            comments = extract_comments_from_faux_title(faux_title)
            nordic_pas_minus_hours = extract_nordic_pas_minus_hours(faux_title)

        person_name_value, is_vikar = extract_person_info(element)
        if data_person:
            person_name_value = data_person
        if not person_name_value:
            person_name_value = extract_person_name(element)

        institution = extract_institution(element, faux_title, institution_overrides)

        for split_row in split_shift_rows(
            start_dt,
            end_dt,
            work_type,
            comments,
            person_name_value,
            institution=institution,
            is_vikar=is_vikar,
            nordic_pas_minus_hours=nordic_pas_minus_hours,
        ):
            rows.append(split_row)

    return rows


def split_shift_rows(
    start_dt: datetime,
    end_dt: datetime,
    work_type: str,
    comments: str,
    person_name: str,
    institution: str = "",
    is_vikar: bool = False,
    nordic_pas_minus_hours: float | None = None,
) -> List[Dict[str, Any]]:
    start_date = start_dt.date()
    end_date = end_dt.date()
    if start_date == end_date:
        single_row = [build_shift_row(
            start_dt,
            end_dt,
            work_type,
            comments,
            person_name,
            institution=institution,
            is_vikar=is_vikar,
            nordic_pas_minus_hours=nordic_pas_minus_hours,
        )]
        has_night_candidate = any(_row_has_night_reduction_candidate(row) for row in single_row)
        night_candidate_rows = sum(1 for row in single_row if _row_has_night_reduction_candidate(row))
        for row in single_row:
            row["event_has_night_reduction_candidate"] = has_night_candidate
            row["event_night_reduction_candidate_rows"] = night_candidate_rows
            row["event_row_count"] = len(single_row)
        return single_row

    rows: List[Dict[str, Any]] = []
    first_end = datetime.combine(start_date, datetime.min.time()) + timedelta(days=1)
    rows.append(build_shift_row(
        start_dt,
        first_end,
        work_type,
        comments,
        person_name,
        institution=institution,
        is_vikar=is_vikar,
        nordic_pas_minus_hours=nordic_pas_minus_hours,
    ))

    current_date = start_date + timedelta(days=1)
    while current_date < end_date:
        next_day = current_date + timedelta(days=1)
        rows.append(build_shift_row(
            datetime.combine(current_date, datetime.min.time()),
            datetime.combine(next_day, datetime.min.time()),
            work_type,
            comments,
            person_name,
            institution=institution,
            is_vikar=is_vikar,
            nordic_pas_minus_hours=nordic_pas_minus_hours,
        ))
        current_date = next_day

    last_start = datetime.combine(end_date, datetime.min.time())
    if end_dt > last_start:
        rows.append(build_shift_row(
            last_start,
            end_dt,
            work_type,
            comments,
            person_name,
            institution=institution,
            is_vikar=is_vikar,
            nordic_pas_minus_hours=nordic_pas_minus_hours,
        ))

    has_night_candidate = any(_row_has_night_reduction_candidate(row) for row in rows)
    night_candidate_rows = sum(1 for row in rows if _row_has_night_reduction_candidate(row))
    for row in rows:
        row["event_has_night_reduction_candidate"] = has_night_candidate
        row["event_night_reduction_candidate_rows"] = night_candidate_rows
        row["event_row_count"] = len(rows)

    return rows


def build_shift_row(
    start_dt: datetime,
    end_dt: datetime,
    work_type: str,
    comments: str,
    person_name: str,
    institution: str = "",
    is_vikar: bool = False,
    nordic_pas_minus_hours: float | None = None,
) -> Dict[str, Any]:
    date = start_dt.strftime("%d-%m-%Y")
    day = DANISH_WEEKDAYS.get(start_dt.strftime("%A"), start_dt.strftime("%A"))
    start_time = start_dt.strftime("%H:%M")
    end_time = end_dt.strftime("%H:%M")
    if end_dt.time() == datetime.min.time():
        end_time = "24:00"
    if start_dt.time() == datetime.min.time():
        start_time = "00:00"

    normalized_work_type = normalize_work_type(work_type)

    return {
        "date": date,
        "day": day,
        "start_time": start_time,
        "end_time": end_time,
        "work_type": normalized_work_type,
        "comments": comments,
        "person": person_name,
        "institution": institution,
        "vikar": is_vikar,
        "nordic_pas_minus_hours": nordic_pas_minus_hours,
    }


def normalize_work_type(work_type: str | None) -> str:
    cleaned = (work_type or "Normtid").strip()
    if not cleaned:
        return ""

    normalized = re.sub(r"\s+", " ", cleaned).strip().lower()
    if normalized in IGNORED_WORK_TYPES:
        return ""

    if normalized in WORK_TYPE_ALIASES:
        return "Normtid"
    if normalized == "fri med telefon":
        return "Normtid"

    mapping = {
        "normtid": "Normtid",
        "ferie": "Ferie",
        "bagvagt": "Bagvagt",
        "batvagt": "Batvagt",
        "+timer": "+timer",
        "+ timer": "+timer",
        "barn 1. sygedag": "Barn 1. Sygedag",
        "barsel": "Barsel",
        "helligdag": "Helligdag",
        "koloni": "Koloni",
        "kursus": "Kursus",
        "kolonitur": "Kolonitur",
        "akademi uddannelse": "Akademi Uddannelse",
        "førstehjælp kursus": "Førstehjælp Kursus",
        "ledelsesseminar": "Ledelsesseminar",
        "lederseminar": "Lederseminar",
        "ledermøde": "Ledermøde",
        "p-dag": "P-Dag",
        "p-møde": "P-Møde",
        "projekt overgreb": "Projekt Overgreb",
        "projekt a": "Projekt A",
        "projekt b": "Projekt B",
        "projekt c": "Projekt C",
        "syg": "Syg",
        "administrativ tid": "Administrativ tid",
        "administrativ tid hjemme": "Administrativ tid hjemme",
        "møde": "Møde",
        "mus-samtale": "MUS-Samtale",
        "udekørende besøg hos lukas": "Udekørende Besøg hos Lukas",
        "udekørende besøg hos sigrid": "Udekørende besøg hos Sigrid",
        "seniordag": "Seniordag",
        "marianne": "Marianne",
        "vikar": "Vikar",
        "vikar stegdalsvej": "Vikar Stegdalsvej",
        "vikar borum": "Vikar Borum",
        "vikar frøstrupskovvej": "Vikar Frøstrupskovvej",
        "vikar fsv2": "Vikar FSV2",
        "alpha vikar fv2": "Alpha Vikar FV2",
        "vikar frøstrupvej1": "Vikar Frøstrupvej1",
        "vikar havndal": "Vikar Havndal",
        "vikar lyngbygade": "Vikar Lyngbygade",
        "vikar vrangstrupvej": "Vikar Vrangstrupvej",
        "vikar ørnereden": "Vikar Ørnereden",
        "vikar - oph": "Vikar - OPH",
        "vikar - aw": "Vikar - AW",
        "vikar feriekolonien": "Vikar Feriekolonien",
        "vikar solo": "Vikar Solo",
        "vikar stn": "Vikar STN",
        "vikar stu": "Vikar STU",
        "vikar paragraf12 stk3": "Vikar Paragraf12 stk3",
        "vikar hesselagergård": "Vikar Hesselagergård",
    }

    return mapping.get(normalized, cleaned)


def normalized_work_type_for_row(work_type: str | None) -> str:
    return normalize_work_type(work_type)


def abbreviate_work_type(work_type: str | None, row: Dict[str, Any] | None = None, apply_night_reduction_to_bagvagt: bool = True) -> str:
    normalized = normalize_work_type(work_type)
    if not normalized:
        return ""

    is_vikar = bool(row and row.get("vikar"))
    if is_vikar:
        if normalized == "Bagvagt":
            if row is not None and _row_has_night_reduction_candidate(row):
                if not apply_night_reduction_to_bagvagt:
                    return "B"
                return "BN"
            return "B"

        if row is not None and is_full_night_shift(row):
            return "NN"
        return "N"

    normalized_key = normalized.lower()
    abbreviation = WORK_TYPE_ABBREVIATIONS.get(normalized_key)
    if abbreviation is None:
        if normalized == "Normtid":
            abbreviation = "N"
        elif normalized == "+timer":
            abbreviation = "P"
        else:
            abbreviation = "N"

    if row is not None and (normalized == "Bagvagt" and _row_has_night_reduction_candidate(row) or normalized != "Bagvagt" and is_full_night_shift(row)):
        if normalized == "Bagvagt" and not apply_night_reduction_to_bagvagt:
            return "B"
        if abbreviation == "N":
            return "NN"
        if abbreviation.endswith("N"):
            return abbreviation
        return f"{abbreviation}N"

    return abbreviation


def is_full_night_shift_span(start_dt: datetime, end_dt: datetime) -> bool:
    start_minutes = start_dt.hour * 60 + start_dt.minute
    if end_dt.time() == datetime.min.time():
        end_minutes = 24 * 60
    else:
        end_minutes = end_dt.hour * 60 + end_dt.minute
    return start_minutes == 0 and end_minutes >= 8 * 60


def is_full_night_shift(row: Dict[str, Any]) -> bool:
    start_time = row.get("start_time", "")
    end_time_value = row.get("end_time", "")
    if end_time_value == "24:00":
        end_minutes = 24 * 60
    else:
        end_minutes = _time_to_minutes(end_time_value)
    return start_time == "00:00" and end_minutes >= 8 * 60


def _time_to_minutes(time_value: str) -> int:
    if not time_value:
        return 0
    hour, minute = map(int, time_value.split(":", 1))
    return hour * 60 + minute


def _row_has_night_reduction_candidate(row: Dict[str, Any]) -> bool:
    start_time = row.get("start_time", "")
    end_time_value = row.get("end_time", "00:00")
    if start_time != "00:00":
        return False
    return _time_to_minutes(end_time_value) >= 6 * 60


def _parse_row_time_range(row: Dict[str, Any]) -> tuple[datetime, datetime]:
    row_date_str = row["date"]
    row_date = None
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            row_date = datetime.strptime(row_date_str, fmt)
            date_format = fmt
            break
        except ValueError:
            continue
    if row_date is None:
        raise ValueError(f"Unsupported date format: {row_date_str}")

    start_time = row.get("start_time", "00:00")
    end_time = row.get("end_time", "00:00")

    if start_time == "24:00":
        start_dt = datetime.combine(row_date.date() + timedelta(days=1), datetime.min.time())
    else:
        start_dt = datetime.strptime(f"{row_date_str} {start_time}", f"{date_format} %H:%M")

    if end_time == "24:00":
        end_dt = datetime.combine(row_date.date() + timedelta(days=1), datetime.min.time())
    else:
        end_dt = datetime.strptime(f"{row_date_str} {end_time}", f"{date_format} %H:%M")

    return start_dt, end_dt


def _nights_window_intersects(start_dt: datetime, end_dt: datetime) -> bool:
    window_start = datetime.combine(start_dt.date(), datetime.min.time())
    window_end = window_start + timedelta(hours=8)
    return start_dt < window_end and end_dt > window_start


def _row_overlaps_night_window(raw_row: Dict[str, Any]) -> bool:
    row_start, row_end = _parse_row_time_range(raw_row)
    return _nights_window_intersects(row_start, row_end)


def extract_work_type(element: Any, faux_title: str) -> str:
    if faux_title:
        title_match = re.search(r"Type:\s*([^<\r\n]+)", faux_title, flags=re.IGNORECASE)
        if title_match:
            return title_match.group(1).strip()

    raw_type = element.get("data-bar") or ""
    if raw_type:
        try:
            payload = json.loads(raw_type)
            if payload.get("type"):
                return str(payload.get("type"))
        except json.JSONDecodeError:
            pass

    for span in element.find_all("span", recursive=False):
        text = span.get_text(strip=True)
        if not text or text in {"", "-"}:
            continue
        normalized = text.strip().lower()
        if normalized in KNOWN_WORK_TYPES:
            return text.strip()
        if normalized in WORK_TYPE_ALIASES:
            return "Normtid"

    return ""


def extract_person_name(element: Any) -> str:
    row = element.find_parent("tr")
    if row is None:
        return ""

    for cell in row.find_all("td", recursive=False):
        text = cell.get_text(" ", strip=True)
        if text:
            return text

    return ""


def extract_person_info(element: Any) -> tuple[str, bool]:
    row = element.find_parent("tr")
    person_name = ""
    is_vikar = False
    if row is not None:
        for cell in row.find_all("td", recursive=False):
            cell_text = cell.get_text(" ", strip=True)
            if not person_name and cell_text:
                person_name = cell_text
            if "opholdssteder - vikarer" in cell_text.lower():
                is_vikar = True
            faux_title = cell.get("faux-title") or ""
            if "opholdssteder - vikarer" in faux_title.lower():
                is_vikar = True

    title = element.get("faux-title") or ""
    if "opholdssteder - vikarer" in title.lower():
        is_vikar = True

    return person_name, is_vikar


def extract_institution(element: Any, faux_title: str, institution_overrides: Dict[str, str] | None = None) -> str:
    institution = ""
    if faux_title:
        title_text = html.unescape(faux_title)
        match = re.search(r"Institution:\s*([^<\r\n]+)", title_text, flags=re.IGNORECASE)
        if match:
            institution = match.group(1).strip()

    if not institution:
        for span in element.find_all("span", recursive=False):
            text = span.get_text(strip=True)
            if not text:
                continue
            normalized_text = text.lower()
            if "alpha.dk" in normalized_text or "opholdssteder" in normalized_text:
                institution = text
                break

    if not institution:
        raw_data = element.get("data-bar") or ""
        if raw_data:
            try:
                payload = json.loads(raw_data)
                if payload.get("inst"):
                    institution = str(payload.get("inst"))
            except json.JSONDecodeError:
                pass

    institution = institution.replace("\u00a0", " ").strip()
    if not institution:
        return ""

    if institution_overrides:
        override = institution_overrides.get(institution)
        if override is None:
            override = institution_overrides.get(institution.lower())
        if override:
            return override

    return INSTITUTION_NAME_OVERRIDES.get(institution, institution)


def extract_comments_from_faux_title(faux_title: str) -> str:
    if not faux_title:
        return ""

    text = html.unescape(faux_title)
    text = text.replace("<br />", "\n").replace("<br>", "\n")
    text = re.sub(r"<[^>]*>", "", text)
    text = text.replace("\u00a0", " ")

    marker = re.search(r"Bemærkninger:\s*(.*)", text, flags=re.IGNORECASE | re.DOTALL)
    if marker:
        comment_text = marker.group(1).strip()
        if comment_text:
            comment_text = re.sub(r"\s*\n\s*", " ", comment_text)
            comment_text = re.sub(r"\s+", " ", comment_text).strip()
            if "Historisk:" in comment_text:
                comment_text = comment_text.split("Historisk:", 1)[0].strip()
            return comment_text

    return ""


def extract_nordic_pas_minus_hours(faux_title: str) -> float | None:
    if not faux_title:
        return None

    text = html.unescape(faux_title)
    text = text.replace("<br />", "\n").replace("<br>", "\n")
    text = re.sub(r"<[^>]*>", "", text)
    text = text.replace("\u00a0", " ")

    # Ignore historical entries; only active/current minus values should count.
    historical_match = re.search(r"\bhistorisk\s*:", text, flags=re.IGNORECASE)
    if historical_match:
        text = text[:historical_match.start()]

    match = re.search(r"minus:\s*([0-9]+(?:[.,][0-9]+)?)", text, flags=re.IGNORECASE)
    if not match:
        return None

    value_text = match.group(1).replace(",", ".")
    try:
        return float(value_text)
    except ValueError:
        return None


def _normalize_nordic_pas_minus_hours(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        value = value.replace(",", ".")
        try:
            return float(value)
        except ValueError:
            return None
    return None


def _night_minus_alignment(row: Dict[str, Any], event_night_reduction_enabled: bool) -> str:
    nordic_pas_minus_hours = _normalize_nordic_pas_minus_hours(row.get("nordic_pas_minus_hours"))
    event_night_candidate_rows = row.get("event_night_reduction_candidate_rows")
    if event_night_candidate_rows is None:
        event_night_candidate_rows = 1 if _row_has_night_reduction_candidate(row) else 0
    current_row_night_candidate = _row_has_night_reduction_candidate(row)
    event_has_night_candidates = bool(row.get("event_has_night_reduction_candidate") or event_night_candidate_rows > 0)

    expected_event_minus_hours = 0.0
    if row.get("event_has_night_reduction_candidate") or event_night_candidate_rows > 0:
        if event_night_reduction_enabled:
            expected_event_minus_hours = float(event_night_candidate_rows) * 2.0

    actual_event_minus_hours = nordic_pas_minus_hours or 0.0
    if nordic_pas_minus_hours is None and expected_event_minus_hours == 0.0:
        return ""

    if actual_event_minus_hours != expected_event_minus_hours:
        if event_has_night_candidates and not current_row_night_candidate:
            return ""
        return "Stemmer ikke"
    return ""


def apply_adjustments(row: Dict[str, Any], apply_night_reduction_to_bagvagt: bool = True) -> Dict[str, Any]:
    start_time = datetime.strptime(row["start_time"], "%H:%M")
    end_time_value = row["end_time"]
    if end_time_value == "24:00":
        end_time_value = "00:00"
        end_time = datetime.strptime(end_time_value, "%H:%M") + timedelta(days=1)
    else:
        end_time = datetime.strptime(end_time_value, "%H:%M")
    duration_hours = (end_time - start_time).total_seconds() / 3600

    original = round(duration_hours, 2)
    adjusted = original
    notes: List[str] = []

    work_type = row.get("work_type", "")
    normalized_work_type = normalized_work_type_for_row(work_type)

    registered_minus_hours = row.get("registered_minus_hours")
    if registered_minus_hours is None:
        registered_minus_hours = row.get("minus_hours")
    registered_minus_hours_source = row.get("registered_minus_hours_source") or ""
    if registered_minus_hours_source == "html":
        registered_minus_hours = None
    registered_minus_hours_display = ""

    automatic_night_reduction_applies = (
        row.get("start_time") == "00:00"
        and _time_to_minutes(row.get("end_time", "00:00")) >= 6 * 60
        and not (normalized_work_type == "Bagvagt" and not apply_night_reduction_to_bagvagt)
    )

    night_deduction_applied = False
    if registered_minus_hours is not None:
        registered_minus_hours_value = float(registered_minus_hours)
        if registered_minus_hours_value == 0:
            registered_minus_hours_display = ""
        else:
            registered_minus_hours_display = f"{int(registered_minus_hours_value)} timer"
            adjusted = round(adjusted - registered_minus_hours_value, 2)
            if registered_minus_hours_source == "nordic_pas":
                if registered_minus_hours_value == 2 and automatic_night_reduction_applies:
                    registered_minus_hours_display = ""
                    night_deduction_applied = True
                else:
                    notes.append(f"Minus timer fra NordicPas: {registered_minus_hours_display}")
            else:
                notes.append(f"Minus timer: {registered_minus_hours_display}")
            if registered_minus_hours_value == 2 and automatic_night_reduction_applies:
                night_deduction_applied = True
    elif automatic_night_reduction_applies:
        adjusted = round(adjusted - 2, 2)
        notes.append("Nat: -2 timer")
        registered_minus_hours_display = "2 timer"
        night_deduction_applied = True

    if normalized_work_type == "Bagvagt":
        adjusted = round(adjusted * 0.5, 2)
        notes.append("Bagvagt: 50% reduktion")

    if normalized_work_type == "":
        adjusted = 0
        notes = []

    if normalized_work_type == "Barn 1. Sygedag" and adjusted > 7.4:
        adjusted = 7.4

    event_night_reduction_enabled = not (
        normalized_work_type == "Bagvagt" and not apply_night_reduction_to_bagvagt
    )
    night_minus_status = _night_minus_alignment(row, event_night_reduction_enabled)

    result = dict(row)
    result["original_duration_hours"] = original
    result["adjusted_duration_hours"] = adjusted
    result["adjustment_notes"] = notes
    result["registered_minus_hours"] = registered_minus_hours_display
    result["registered_minus_hours_source"] = registered_minus_hours_source
    result["expected_night_reduction"] = 2 if automatic_night_reduction_applies and registered_minus_hours is None else 0
    result["night_deduction"] = night_deduction_applied
    result["night_minus_status"] = night_minus_status
    return result


def export_to_excel(
    rows: List[Dict[str, Any]],
    template_path: str | Path,
    output_path: str | Path | None = None,
    apply_night_reduction_to_bagvagt: bool = True,
    pay_period_mode: str = "16-15",
) -> Path:
    template = Path(template_path)
    if output_path is None:
        output_path = template

    output = Path(output_path)
    workbook = Workbook()
    workbook.remove(workbook.active)

    headers = [
        "Dato",
        "Dag",
        "Starttid",
        "Sluttid",
        "Vagtform",
        "Institution",
        "Personsnavn",
        "Vikar",
        "Summering efter reduktioner",
        "Fulde vagtforms navn",
        "Bemærkning",
        "Justeringsnote",
        "Summering før reduktioner",
        "Registreret minus timer",
        "Afstemning",
    ]
    header_last_col = get_column_letter(len(headers))

    if not rows:
        if "Data" in workbook.sheetnames:
            workbook.remove(workbook["Data"])
        sheet = workbook.create_sheet(title="Data")
        sheet.append(headers)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{header_last_col}1"
        sheet.append(["Ingen vagter fundet i den angivne HTML."])
        target_path = _persist_workbook(workbook, output)
        return target_path

    deduped_rows = deduplicate_rows(rows)
    filtered_rows = [row for row in deduped_rows if is_within_pay_period(row, deduped_rows, pay_period_mode=pay_period_mode)]
    if not filtered_rows:
        sheet = workbook.create_sheet(title="Data")
        sheet.append(headers)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{header_last_col}1"
        sheet.append(["Ingen vagter fundet i den angivne HTML."])
        return _persist_workbook(workbook, output)
    for row in filtered_rows:
        if normalize_work_type(row.get("work_type")) == "":
            continue
        adjusted = apply_adjustments(row, apply_night_reduction_to_bagvagt=apply_night_reduction_to_bagvagt)
        if adjusted["adjusted_duration_hours"] == 0 and not adjusted["adjustment_notes"]:
            continue
        sheet_name = safe_sheet_name(adjusted["person"])
        if sheet_name not in workbook.sheetnames:
            sheet = workbook.create_sheet(title=sheet_name)
            sheet.append(headers)
            sheet.freeze_panes = "A2"
        else:
            sheet = workbook[sheet_name]

        display_work_type = abbreviate_work_type(
            adjusted.get("work_type"),
            adjusted,
            apply_night_reduction_to_bagvagt=apply_night_reduction_to_bagvagt,
        )
        adjusted = dict(adjusted)
        adjusted["work_type"] = display_work_type
        if adjusted["work_type"] == "":
            continue
        original_work_type = row.get("work_type", "")
        sheet.append([
            adjusted["date"],
            adjusted["day"],
            adjusted["start_time"],
            adjusted["end_time"],
            adjusted["work_type"],
            adjusted.get("institution", ""),
            adjusted["person"],
            "Vikar" if adjusted.get("vikar") else "",
            adjusted["adjusted_duration_hours"],
            original_work_type,
            adjusted["comments"],
            "; ".join(adjusted["adjustment_notes"]),
            adjusted["original_duration_hours"],
            adjusted.get("registered_minus_hours", ""),
            adjusted.get("night_minus_status", ""),
        ])

        last_row = sheet.max_row
        row_fill = None
        if adjusted.get("night_deduction"):
            row_fill = PatternFill(fill_type="solid", fgColor="FFCDD2", bgColor="FFCDD2")
        elif adjusted["work_type"] == "B":
            row_fill = PatternFill(fill_type="solid", fgColor="CFE8FF", bgColor="CFE8FF")
        elif adjusted["work_type"] == "P":
            row_fill = PatternFill(fill_type="solid", fgColor="C6EFC6", bgColor="C6EFC6")
        elif adjusted["work_type"] == "F":
            row_fill = PatternFill(fill_type="solid", fgColor="F5E6CC", bgColor="F5E6CC")
        elif adjusted["work_type"] == "O":
            row_fill = PatternFill(fill_type="solid", fgColor="4F81BD", bgColor="4F81BD")
        elif adjusted["work_type"] == "S":
            row_fill = PatternFill(fill_type="solid", fgColor="CC6666", bgColor="CC6666")
        for cell in sheet[last_row]:
            if adjusted["adjustment_notes"]:
                cell.font = Font(bold=True)
            if row_fill is not None:
                cell.fill = row_fill

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        if sheet.max_row >= 1:
            sheet.freeze_panes = "A2"
        if sheet_name == "Data":
            continue
        if sheet.max_row >= 1:
            filter_end_row = max(1, sheet.max_row)
            sheet.auto_filter.ref = f"A1:{header_last_col}{filter_end_row}"
        if sheet.max_row > 1:
            data_row_count = sheet.max_row - 1
            total_row = sheet.max_row + 1
            sheet.cell(row=total_row, column=1, value="Total")
            sheet.cell(row=total_row, column=2, value="")
            sheet.cell(row=total_row, column=3, value="")
            sheet.cell(row=total_row, column=4, value="")
            sheet.cell(row=total_row, column=5, value="")
            sheet.cell(row=total_row, column=6, value="")
            sheet.cell(row=total_row, column=7, value="")
            sheet.cell(row=total_row, column=8, value="")
            sheet.cell(row=total_row, column=9, value=f"=SUM(I2:I{data_row_count + 1})")
            sheet.cell(row=total_row, column=10, value="")
            sheet.cell(row=total_row, column=11, value="")
            sheet.cell(row=total_row, column=12, value="")
            sheet.cell(row=total_row, column=13, value=f"=SUM(M2:M{data_row_count + 1})")
            sheet.cell(row=total_row, column=14, value="")
            sheet.cell(row=total_row, column=15, value="")
            for cell in sheet[total_row]:
                cell.font = Font(bold=True)

    return _persist_workbook(workbook, output)


def deduplicate_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    deduped: List[Dict[str, Any]] = []
    kept_by_key: Dict[tuple[str, str], Dict[str, Any]] = {}
    kept_indices: Dict[tuple[str, str], int] = {}

    for row in rows:
        key = (row.get("person", ""), row.get("date", ""))
        last_kept = kept_by_key.get(key)
        if last_kept is None:
            kept_by_key[key] = row
            kept_indices[key] = len(deduped)
            deduped.append(row)
            continue

        if _rows_overlap(last_kept, row):
            if _row_is_broader(row, last_kept):
                deduped[kept_indices[key]] = row
                kept_by_key[key] = row
            continue

        kept_by_key[key] = row
        kept_indices[key] = len(deduped)
        deduped.append(row)

    return deduped


def _rows_overlap(first: Dict[str, Any], second: Dict[str, Any]) -> bool:
    first_start = _time_to_minutes(first.get("start_time", "00:00"))
    first_end = _time_to_minutes(first.get("end_time", "00:00"))
    second_start = _time_to_minutes(second.get("start_time", "00:00"))
    second_end = _time_to_minutes(second.get("end_time", "00:00"))
    if first_end == 0:
        first_end = 24 * 60
    if second_end == 0:
        second_end = 24 * 60
    return max(first_start, second_start) < min(first_end, second_end)


def _row_is_broader(candidate: Dict[str, Any], current: Dict[str, Any]) -> bool:
    candidate_start = _time_to_minutes(candidate.get("start_time", "00:00"))
    candidate_end = _time_to_minutes(candidate.get("end_time", "00:00"))
    current_start = _time_to_minutes(current.get("start_time", "00:00"))
    current_end = _time_to_minutes(current.get("end_time", "00:00"))
    if candidate_end == 0:
        candidate_end = 24 * 60
    if current_end == 0:
        current_end = 24 * 60
    return candidate_start <= current_start and candidate_end >= current_end


def _persist_workbook(workbook: Workbook, output: Path) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    target_path = output
    if output.exists():
        try:
            output.unlink()
        except PermissionError:
            pass

    if output.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = output.with_name(f"{output.stem}_{timestamp}{output.suffix}")

    try:
        workbook.save(target_path)
        return target_path
    except PermissionError:
        counter = 1
        fallback_path = target_path
        while fallback_path.exists():
            fallback_path = target_path.with_name(f"{target_path.stem}_{counter}{target_path.suffix}")
            counter += 1
        workbook.save(fallback_path)
        return fallback_path


def is_within_pay_period(row: Dict[str, Any], rows: List[Dict[str, Any]] | None = None, pay_period_mode: str = "15-16") -> bool:
    try:
        row_date = datetime.strptime(row.get("date", ""), "%d-%m-%Y")
    except ValueError:
        return True

    if rows:
        parsed_dates = []
        for candidate in rows:
            try:
                parsed_dates.append(datetime.strptime(candidate.get("date", ""), "%d-%m-%Y"))
            except ValueError:
                continue
        if parsed_dates:
            first_date = min(parsed_dates)
            if pay_period_mode == "01 - Sidste i måneden":
                start_date = first_date.replace(day=1)
                end_date = (start_date.replace(year=start_date.year + (start_date.month == 12 and 1 or 0), month=start_date.month % 12 + 1) - timedelta(days=1))
                return start_date <= row_date <= end_date

            if pay_period_mode in {"15-16", "16-15"}:
                start_date = first_date.replace(day=16)
                end_date = (start_date.replace(year=start_date.year + (start_date.month == 12 and 1 or 0), month=start_date.month % 12 + 1) - timedelta(days=1))
                if end_date.day != 15:
                    end_date = end_date.replace(day=15)
                return start_date <= row_date <= end_date

            return True

    return True


def safe_sheet_name(person_name: str) -> str:
    sanitized = re.sub(r"[^0-9A-Za-zÆØÅæøå]+", "_", person_name).strip("_") or "Person"
    return sanitized[:31]
