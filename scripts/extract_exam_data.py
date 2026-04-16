#!/usr/bin/env python3
"""Extract normalized exam data from the provided XLSX timetable.

The script avoids third-party dependencies by parsing the XLSX XML payload directly.
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List
import xml.etree.ElementTree as ET

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

LEVEL_ORDER = {"IG": 0, "AS": 1, "A2": 2}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract exam timetable JSON from XLSX")
    parser.add_argument("--input", required=True, help="Path to source XLSX")
    parser.add_argument("--output", required=True, help="Path to output JSON")
    return parser.parse_args()


def parse_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    values: List[str] = []
    for item in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in item.findall(".//a:t", NS)))
    return values


def read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str = "xl/worksheets/sheet1.xml") -> List[Dict[str, str]]:
    shared = parse_shared_strings(zf)
    sheet = ET.fromstring(zf.read(sheet_path))
    sheet_data = sheet.find("a:sheetData", NS)
    if sheet_data is None:
        return []

    rows: List[Dict[str, str]] = []
    for row in sheet_data.findall("a:row", NS):
        payload: Dict[str, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            col_match = re.match(r"[A-Z]+", ref)
            if not col_match:
                continue
            col = col_match.group(0)
            cell_type = cell.attrib.get("t")
            value_node = cell.find("a:v", NS)
            raw_value = "" if value_node is None else (value_node.text or "")

            if cell_type == "s" and raw_value:
                value = shared[int(raw_value)]
            elif cell_type == "b":
                value = "TRUE" if raw_value == "1" else "FALSE"
            elif cell_type == "inlineStr":
                inline = cell.find("a:is", NS)
                value = "".join(node.text or "" for node in inline.findall(".//a:t", NS)) if inline is not None else ""
            else:
                value = raw_value
            payload[col] = value
        rows.append(payload)
    return rows


def excel_date(value: str) -> str:
    value = clean_whitespace(value)
    if not value:
        return ""
    if re.match(r"^\d{4}-\d{2}-\d{2}$", value):
        return value
    dt = datetime(1899, 12, 30) + timedelta(days=float(value))
    return dt.date().isoformat()


def excel_time(value: str) -> str:
    value = clean_whitespace(value)
    if not value:
        return ""
    time_match = re.match(r"^(\d{1,2}):(\d{2})(?::\d{2})?$", value)
    if time_match:
        hours = int(time_match.group(1))
        minutes = int(time_match.group(2))
        return f"{hours:02d}:{minutes:02d}"
    total_seconds = int(round(float(value) * 24 * 60 * 60))
    hours = (total_seconds // 3600) % 24
    minutes = (total_seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}"


def duration_to_minutes(text: str) -> int:
    text = (text or "").strip().lower()
    hours = 0
    minutes = 0
    h_match = re.search(r"(\d+)\s*h", text)
    m_match = re.search(r"(\d+)\s*m", text)
    if h_match:
        hours = int(h_match.group(1))
    if m_match:
        minutes = int(m_match.group(1))
    return hours * 60 + minutes


def clean_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def detect_session_tag(code: str, title: str, syllabus: str) -> str:
    for field in (code, title, syllabus):
        match = re.search(r"-(?:S|Session)\s*(\d+)\b", field or "", flags=re.IGNORECASE)
        if match:
            return f"S{match.group(1)}"
    return "S1"


def strip_session_suffix(text: str) -> str:
    value = re.sub(r"\s*-(?:S|Session)\s*\d+\b", "", text or "", flags=re.IGNORECASE)
    return clean_whitespace(value)


def normalize_subject(syllabus: str) -> str:
    value = re.sub(r"-(?:S|Session)\s*\d+\b", "", syllabus or "", flags=re.IGNORECASE)
    return clean_whitespace(value)


def infer_board(qual_level: str) -> str:
    if "Cambridge" in qual_level:
        return "CIE"
    if "Edexcel" in qual_level:
        return "Edexcel"
    return "Other"


def infer_level(qual_level: str, code: str, title: str) -> str:
    q = qual_level or ""
    if "IGCSE" in q:
        return "IG"

    if "AS & A Level" in q:
        component = code.split("/")[-1].split("-")[0] if "/" in code else ""
        paper = component[0] if component else ""
        subject_code = code.split("/")[0] if "/" in code else ""
        title_upper = (title or "").upper()

        if subject_code == "9868":
            return "AS"

        if subject_code in {"9700", "9701", "9702"}:
            if paper in {"1", "2", "3"}:
                return "AS"
            if paper in {"4", "5"}:
                return "A2"

        if subject_code == "9709":
            if "PURE MATHEMATICS 1" in title_upper or "PROBABILITY & STATISTICS 1" in title_upper:
                return "AS"
            if "PURE MATHEMATICS 3" in title_upper or "MECHANICS" in title_upper:
                return "A2"

        if subject_code == "9231":
            if "FURTHER PURE MATHEMATICS 14" in title_upper or "PROBABILITY & STATISTICS" in title_upper:
                return "AS"
            if "FURTHER PURE MATHEMATICS 24" in title_upper or "MECHANICS" in title_upper:
                return "A2"

        if paper in {"1", "2"}:
            return "AS"
        if paper in {"3", "4", "5", "6", "7", "8", "9"}:
            return "A2"
        return "AS"

    if "IAL" in q:
        token = code.split()[0].upper() if code else ""
        title_upper = (title or "").upper()

        # User-confirmed Edexcel mappings.
        if token in {"WEC11", "WEC12"}:
            return "AS"
        if token in {"WEC13", "WEC14"}:
            return "A2"
        if token in {"WMA11", "WMA12", "WST01"}:
            return "AS"
        if token in {"WMA13", "WMA14", "WME01", "WDM11"}:
            return "A2"

        # General module fallback.
        if any(marker in title_upper for marker in ["P1", "P2", "S1", "UNIT 1", "UNIT 2"]):
            return "AS"
        if any(
            marker in title_upper
            for marker in ["P3", "P4", "M1", "M2", "M3", "S2", "S3", "D1", "D2", "UNIT 3", "UNIT 4", "UNIT 5", "UNIT 6"]
        ):
            return "A2"
        if any(marker in title_upper for marker in ["FP1", "FP2"]):
            return "AS"
        if "FP3" in title_upper:
            return "A2"

        unit_match = re.search(r"UNIT\s*(\d+)", title_upper)
        if unit_match:
            unit_num = int(unit_match.group(1))
            if unit_num <= 2:
                return "AS"
            return "A2"

        return "AS"

    if "GCE A Level" in q:
        return "A2"

    return "AS"


def infer_subject(qual_level: str, syllabus: str, code: str, title: str) -> str:
    subject = normalize_subject(syllabus)
    q = qual_level or ""
    token = code.split()[0].upper() if code else ""
    title_upper = (title or "").upper()

    if "IAL" in q and subject == "Mathematics":
        if token in {"WFM01", "WFM02", "WFM03", "WME02", "WME03", "WST02", "WST03"}:
            return "Further Mathematics"
        if any(marker in title_upper for marker in ["FP1", "FP2", "FP3", "M2", "M3", "S2", "S3"]):
            return "Further Mathematics"
        return "Mathematics"

    return subject


def to_minutes(hhmm: str) -> int:
    if not hhmm:
        return 0
    hours, minutes = hhmm.split(":")
    return int(hours) * 60 + int(minutes)


def infer_timetable_name(rows: List[Dict[str, str]]) -> str:
    series_values = [clean_whitespace(row.get("A", "")) for row in rows[1:] if clean_whitespace(row.get("A", ""))]
    if not series_values:
        return "June 2026 External Exam Timetable"

    most_common, _ = Counter(series_values).most_common(1)[0]
    try:
        series_date = datetime.fromisoformat(excel_date(most_common))
    except ValueError:
        return "June 2026 External Exam Timetable"

    return f"{series_date.strftime('%B %Y')} External Exam Timetable"


def serialize_rows(rows: List[Dict[str, str]]) -> Dict[str, object]:
    families: Dict[str, Dict[str, object]] = {}

    for row in rows[1:]:
        qual_level = clean_whitespace(row.get("B", ""))
        code = clean_whitespace(row.get("C", ""))
        syllabus = clean_whitespace(row.get("D", ""))
        title = clean_whitespace(row.get("E", ""))
        date_value = excel_date(row.get("F", ""))
        start = excel_time(row.get("I", ""))
        end = excel_time(row.get("J", ""))

        if not qual_level or not code or not date_value or not start or not end:
            continue

        session = detect_session_tag(code, title, syllabus)
        base_code = strip_session_suffix(code)
        base_title = strip_session_suffix(title)
        subject = infer_subject(qual_level, syllabus, code, title)
        level = infer_level(qual_level, code, title)
        board = infer_board(qual_level)

        key = f"{qual_level}|{base_code}"
        if key not in families:
            family_id = re.sub(r"[^a-z0-9]+", "-", f"{board}-{level}-{base_code}-{subject}".lower()).strip("-")
            families[key] = {
                "id": family_id,
                "level": level,
                "board": board,
                "qualification": qual_level,
                "subject": subject,
                "code": base_code,
                "paper": base_title,
                "variants": [],
            }

        variant = {
            "session": session,
            "date": date_value,
            "start": start,
            "end": end,
            "startMinutes": to_minutes(start),
            "endMinutes": to_minutes(end),
            "durationMinutes": duration_to_minutes(row.get("H", "")),
            "entries": int(float(row.get("M", "0") or "0")),
            "fullSupervisionRequired": (row.get("K", "") == "TRUE"),
            "arrival": excel_time(row.get("L", "")),
            "venue": clean_whitespace(row.get("N", "")),
            "note": clean_whitespace(row.get("O", "")),
            "originalCode": code,
            "originalTitle": title,
        }

        family = families[key]
        seen_keys = {
            f"{item['session']}|{item['date']}|{item['start']}|{item['end']}" for item in family["variants"]
        }
        variant_key = f"{variant['session']}|{variant['date']}|{variant['start']}|{variant['end']}"
        if variant_key not in seen_keys:
            family["variants"].append(variant)

    data = list(families.values())

    def session_rank(session_value: str) -> int:
        match = re.search(r"(\d+)", session_value)
        return int(match.group(1)) if match else 1

    for family in data:
        family["variants"].sort(key=lambda item: (item["date"], session_rank(item["session"]), item["startMinutes"]))
        family["defaultSession"] = family["variants"][0]["session"] if family["variants"] else "S1"

    data.sort(
        key=lambda item: (
            LEVEL_ORDER.get(item["level"], 99),
            item["variants"][0]["date"] if item["variants"] else "9999-12-31",
            item["variants"][0]["startMinutes"] if item["variants"] else 0,
            item["subject"],
            item["code"],
        )
    )

    all_dates = [variant["date"] for family in data for variant in family["variants"]]
    unique_quals = sorted({family["qualification"] for family in data})

    return {
        "meta": {
            "name": infer_timetable_name(rows),
            "generatedAt": datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
            "levels": ["IG", "AS", "A2"],
            "qualifications": unique_quals,
            "totalExamFamilies": len(data),
            "dateRange": {
                "start": min(all_dates) if all_dates else None,
                "end": max(all_dates) if all_dates else None,
            },
            "timezone": "Asia/Shanghai",
        },
        "examFamilies": data,
    }


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    with zipfile.ZipFile(input_path) as zf:
        rows = read_sheet_rows(zf)

    payload = serialize_rows(rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote {payload['meta']['totalExamFamilies']} exam families to {output_path}")


if __name__ == "__main__":
    main()
