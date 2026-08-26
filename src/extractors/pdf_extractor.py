"""
pdf_extractor.py

PDFExtractor — implements docs/05_pdf_extractor_design.md.

Extracts fleet-forecast data from `PrevisionFlota.pdf` and produces
two structured datasets, exactly as scoped in the design doc:

    - CommercialServicesDataset (per commercial circulation)
    - ReserveDataset (per reserve unit)

Field notes vs. the design doc's minimal schema:
    - CommercialRecord carries `route` in addition to Service and
      Registration. `route` is not in the design doc's required-fields
      table, but it costs nothing to extract (it's already on the
      source line) and keeps the door open for a route cross-check in
      the Validation Layer without having to re-touch this extractor
      later. Downstream layers are free to ignore it.
    - `raw_line` is kept purely for error traceability in logs/reports
      — it is not a business field and is not part of either dataset's
      required schema.

Why word-position grouping instead of pdfplumber's `extract_tables()`:
the source PDF has no real table structure (no cell borders to key
off of) — it's positioned text. `extract_tables()` was tested against
the actual file and produced garbage (merged/fragmented rows). Rows
are instead reconstructed by grouping words with the same vertical
("top") coordinate and sorting each group left-to-right by "x0".

Per docs/01_requirements.md (scope) and 05_pdf_extractor_design.md
(out of scope), maintenance and non-productive-train blocks are read
but discarded — they are not part of either output dataset.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field

import pdfplumber

REGISTRATION_RE = re.compile(r"^[A-Z]\d{3}$")
SERVICE_RE = re.compile(r"^\d{3}(-\d{3})?$")
DATE_RE = re.compile(r"^\d{2}-[a-zé]{3}-\d{2}$", re.IGNORECASE)
ROW_MERGE_TOLERANCE_PX = 3  # max vertical gap (px) to treat words as same line


@dataclass
class CommercialRecord:
    """
    One row of the CommercialServicesDataset.

    Required by 05_pdf_extractor_design.md: service, registration.
    Extra (see module docstring): route, raw_line.
    """

    service: str  # e.g. "101" or "301-302" (not yet split — Transformation Layer's job)
    registration: str
    route: str
    raw_line: str


@dataclass
class ReserveRecord:
    """One row of the ReserveDataset — matches the design doc's schema exactly."""

    workshop_station: str
    registration: str
    status: str  # "RESERVA" or "RESERVA EN ESTACION"
    raw_line: str


@dataclass
class PDFExtractionResult:
    commercial_services: list[CommercialRecord] = field(default_factory=list)
    reserve: list[ReserveRecord] = field(default_factory=list)
    skipped_lines: list[str] = field(default_factory=list)  # maintenance, non-productive, headers


def _group_words_into_lines(words: list[dict]) -> list[list[dict]]:
    """Group words with the same (or very close) `top` into ordered lines."""
    rows_by_top: dict[int, list[dict]] = defaultdict(list)
    for w in words:
        rows_by_top[round(w["top"])].append(w)

    sorted_tops = sorted(rows_by_top.keys())
    if not sorted_tops:
        return []

    merged_groups: list[list[int]] = [[sorted_tops[0]]]
    for t in sorted_tops[1:]:
        if t - merged_groups[-1][-1] <= ROW_MERGE_TOLERANCE_PX:
            merged_groups[-1].append(t)
        else:
            merged_groups.append([t])

    lines = []
    for group in merged_groups:
        line_words = []
        for t in group:
            line_words.extend(rows_by_top[t])
        line_words.sort(key=lambda w: w["x0"])
        lines.append(line_words)
    return lines


def _tokens(line_words: list[dict]) -> list[str]:
    return [w["text"] for w in line_words]

def _find_shared_registrations(
    token_lines: list[list[str]],
) -> dict[int, str]:
    """Associate a vertically centred registration with adjacent service lines."""
    shared_registrations: dict[int, str] = {}

    for center_idx, tokens in enumerate(token_lines):
        if len(tokens) < 3:
            continue

        if not tokens[1].startswith("TS"):
            continue

        registration = tokens[2]
        if not REGISTRATION_RE.match(registration):
            continue

        # A normal commercial row already contains its own service and
        # must continue through the standard parser.
        if any(SERVICE_RE.match(token) for token in tokens):
            continue

        for neighbor_idx in (center_idx - 1, center_idx + 1):
            if not 0 <= neighbor_idx < len(token_lines):
                continue

            neighbor_tokens = token_lines[neighbor_idx]
            has_service = any(
                SERVICE_RE.match(token) for token in neighbor_tokens
            )
            has_registration = any(
                REGISTRATION_RE.match(token) for token in neighbor_tokens
            )

            if has_service and not has_registration:
                shared_registrations[neighbor_idx] = registration

    return shared_registrations

def _parse_commercial_line(
    tokens: list[str],
    raw_line: str,
    shared_registration: str | None = None,
) -> CommercialRecord | None:
    """
    Parse a standard commercial line or a service line whose registration
    is stored in a vertically centred adjacent row.
    """
    if shared_registration is None:
        if len(tokens) < 4:
            return None

        if not REGISTRATION_RE.match(tokens[2]):
            return None

        registration = tokens[2]
        service_search_start = 3
    else:
        if not REGISTRATION_RE.match(shared_registration):
            return None

        registration = shared_registration
        service_search_start = 0

    service_idx = None
    for i in range(service_search_start, len(tokens)):
        if SERVICE_RE.match(tokens[i]):
            service_idx = i
            break

    if service_idx is None:
        return None

    service = tokens[service_idx]
    route = tokens[service_idx + 1] if service_idx + 1 < len(tokens) else ""

    return CommercialRecord(
        service=service,
        registration=registration,
        route=route,
        raw_line=raw_line,
    )


def _parse_reserve_line(tokens: list[str], raw_line: str) -> ReserveRecord | None:
    """
    Expected token shape:
    Station TS-code Registration '-' [time] RESERVA [EN ESTACION]

    Example tokens:
    ['Cancún', 'TS08', 'D008', '-', '12h00', 'RESERVA', 'EN', 'ESTACION']
    ['Cancún', 'TS19', 'N002', '-', 'RESERVA']
    """
    if "RESERVA" not in tokens:
        return None
    if len(tokens) < 4:
        return None
    if not REGISTRATION_RE.match(tokens[2]):
        return None

    reserva_idx = tokens.index("RESERVA")
    status = " ".join(tokens[reserva_idx:])

    return ReserveRecord(
        workshop_station=tokens[0],
        registration=tokens[2],
        status=status,
        raw_line=raw_line,
    )

def _parse_useful_commercial_line(
    tokens: list[str],
    raw_line: str,
) -> ReserveRecord | None:
    """Parse an available train as a reserve with a blank output status."""
    if len(tokens) < 4:
        return None

    ts_idx = next(
        (
            index
            for index, token in enumerate(tokens)
            if re.match(r"^TS\d+$", token)
        ),
        None,
    )
    if ts_idx is None or ts_idx + 2 >= len(tokens):
        return None

    station = " ".join(tokens[:ts_idx]).strip()
    registration = tokens[ts_idx + 1]
    availability_marker = tokens[ts_idx + 2]

    if not station:
        return None

    if not REGISTRATION_RE.match(registration):
        return None

    # Useful trains show '-' after the registration. Rows with a date
    # belong to maintenance or other non-available conditions.
    if availability_marker != "-":
        return None

    return ReserveRecord(
        workshop_station=station,
        registration=registration,
        status="RESERVA",
        raw_line=raw_line,
    )


def _is_maintenance_or_non_productive(tokens: list[str]) -> bool:
    """
    Maintenance rows have a date token like '01-jul-26' in position 3
    (instead of '-'). Non-productive rows are handled separately via
    the section-header flag in extract() — both are out of scope per
    01_requirements.md and 05_pdf_extractor_design.md.
    """
    if len(tokens) >= 4 and DATE_RE.match(tokens[3]):
        return True
    return False


class PDFExtractor:
    """
    Implements docs/05_pdf_extractor_design.md.

    Extraction only — no validation, no business rules, no Excel
    writes (all explicitly out of scope per the design doc).
    """

    def extract(self, pdf_path: str) -> PDFExtractionResult:
        result = PDFExtractionResult()
        in_non_productive_section = False
        in_useful_commercial_section = False

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
                lines = _group_words_into_lines(words)
                token_lines = [_tokens(line_words) for line_words in lines]
                shared_registrations = _find_shared_registrations(token_lines)

                for line_idx, tokens in enumerate(token_lines):
                    raw_line = " ".join(tokens)

                    if not tokens:
                        continue

                    if (
                        tokens[0] == "TRENES"
                        and "UTILES" in tokens
                        and "COMERCIAL" in tokens
                    ):
                        in_useful_commercial_section = True
                        result.skipped_lines.append(raw_line)
                        continue

                    if tokens[0] == "TRENES" and "PRODUCTIVOS" in tokens:
                        in_useful_commercial_section = False
                        in_non_productive_section = True
                        result.skipped_lines.append(raw_line)
                        continue

                    if in_non_productive_section:
                        result.skipped_lines.append(raw_line)
                        continue

                    if in_useful_commercial_section:
                        useful_record = _parse_useful_commercial_line(
                            tokens,
                            raw_line,
                        )
                        if useful_record:
                            result.reserve.append(useful_record)
                        else:
                            result.skipped_lines.append(raw_line)
                        continue

                    if _is_maintenance_or_non_productive(tokens):
                        result.skipped_lines.append(raw_line)
                        continue

                    reserve_record = _parse_reserve_line(tokens, raw_line)
                    if reserve_record:
                        result.reserve.append(reserve_record)
                        continue

                    commercial_record = _parse_commercial_line(
                        tokens, 
                        raw_line,
                        shared_registrations.get(line_idx),
                    )
                    if commercial_record:
                        result.commercial_services.append(commercial_record)
                        continue

                    # Header lines, page titles, etc. — not data.
                    result.skipped_lines.append(raw_line)

        return result


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "data/raw/prevision_flota_sample.pdf"
    extractor = PDFExtractor()
    res = extractor.extract(path)

    print(f"CommercialServicesDataset: {len(res.commercial_services)} records")
    for rec in res.commercial_services:
        print(f"  service={rec.service:10s} registration={rec.registration}  route={rec.route}")

    print(f"\nReserveDataset: {len(res.reserve)} records")
    for rec in res.reserve:
        print(f"  workshop_station={rec.workshop_station:12s} registration={rec.registration}  status={rec.status}")

    print(f"\nSkipped lines (maintenance/non-productive/headers): {len(res.skipped_lines)}")
