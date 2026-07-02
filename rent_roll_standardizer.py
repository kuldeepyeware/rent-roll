#!/usr/bin/env python3
"""
Rent Roll Standardizer v0

Deterministic spreadsheet structure parsing with narrowly-scoped OpenRouter
assistance for uncertain column and charge semantics.

Runtime dependencies:
    pip install openpyxl xlrd

CSV files use only the Python standard library. openpyxl is required for .xlsx
and xlrd is required for legacy .xls workbooks.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Optional


VERSION = "0.1.0"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
DEFAULT_MODEL = "anthropic/claude-haiku-4.5"
MONEY_QUANTUM = Decimal("0.01")

CANONICAL_FIELDS = {
    "unit_number",
    "unit_type",
    "unit_sqft",
    "tenant_name",
    "status",
    "market_rent",
    "effective_rent",
    "total_rent",
    "move_in_date",
    "lease_start_date",
    "lease_end_date",
    "move_out_date",
    "balance",
    "deposit",
    "charge_code",
    "charge_description",
    "charge_amount",
    "property_name",
    "rent_roll_date",
    "ignore",
    "unknown",
}

FIELD_ALIASES = {
    "unit_number": [
        "unit", "unit no", "unit number", "unit #", "apt", "apt no",
        "apartment", "bldg unit", "building unit", "space", "space no",
    ],
    "unit_type": [
        "unit type", "type", "floorplan", "floor plan", "bed bath", "beds baths",
        "model", "unit model",
    ],
    "unit_sqft": [
        "sqft", "sq ft", "square feet", "unit size", "size", "area", "sf",
    ],
    "tenant_name": [
        "tenant", "tenant name", "resident", "resident name", "occupant",
        "lessee", "customer", "name",
    ],
    "status": [
        "status", "unit status", "occupancy", "occupancy status", "lease status",
    ],
    "market_rent": [
        "market rent", "market", "asking rent", "gross market rent", "gpr",
    ],
    "effective_rent": [
        "effective rent", "net rent", "net effective rent", "actual rent",
    ],
    "total_rent": [
        "total rent", "monthly rent", "lease rent", "current rent", "scheduled rent",
    ],
    "move_in_date": [
        "move in", "move in date", "move-in", "move-in date", "mi date",
    ],
    "lease_start_date": [
        "lease start", "lease start date", "lease begin", "lease from",
        "commencement", "start date",
    ],
    "lease_end_date": [
        "lease end", "lease end date", "lease expiration", "lease expiry",
        "expiration", "expiry", "end date", "lease to",
    ],
    "move_out_date": [
        "move out", "move out date", "move-out", "move-out date", "mo date",
    ],
    "balance": [
        "balance", "resident balance", "tenant balance", "amount due", "receivable",
    ],
    "deposit": [
        "deposit", "security deposit", "sec deposit", "deposit held",
    ],
    "charge_code": [
        "charge code", "charge", "code", "transaction code", "rent code",
    ],
    "charge_description": [
        "charge description", "description", "charge name", "charge type",
    ],
    "charge_amount": [
        "charge amount", "amount", "monthly charge", "charge value",
    ],
    "property_name": ["property", "property name", "community"],
    "rent_roll_date": ["rent roll date", "as of", "as of date", "report date"],
}

CHARGE_ALIASES = {
    "BASE_RENT": [
        "rent", "base rent", "monthly rent", "apartment rent", "rnt", "rent charge",
    ],
    "CONCESSION": [
        "concession", "concessions", "conc", "discount", "rent credit",
        "lease incentive", "free rent", "empdisc",
    ],
    "UTILITY": [
        "utility", "utilities", "water", "sewer", "gas", "electric", "electricity",
        "rub", "rubs",
    ],
    "PARKING": ["parking", "garage", "carport", "park"],
    "PET_FEE": ["pet", "pet fee", "pet rent", "petrent", "animal"],
    "PEST_FEE": ["pest", "pest fee", "pest control"],
    "TRASH_FEE": ["trash", "trash fee", "garbage", "valet trash"],
    "STORAGE": ["storage", "locker"],
    "ADMIN_FEE": ["admin", "admin fee", "administrative fee", "billing fee"],
    "OTHER_RENT": [
        "amenity", "amenity fee", "premium", "unit premium", "washer dryer",
        "w d", "cable", "internet", "technology",
    ],
    "ONE_TIME_FEE": [
        "application fee", "late fee", "nsf", "damage", "cleaning fee",
        "move in fee", "termination fee",
    ],
    "DEPOSIT": ["deposit", "security deposit", "pet deposit"],
}

RECURRING_CHARGE_CATEGORIES = {
    "BASE_RENT", "UTILITY", "PARKING", "PET_FEE", "PEST_FEE", "TRASH_FEE",
    "STORAGE", "ADMIN_FEE", "OTHER_RENT",
}

OCCUPIED_WORDS = {
    "occupied", "current", "leased", "resident", "notice", "ntv",
    "month to month", "mtm",
}
VACANT_WORDS = {
    "vacant", "vac", "available", "ready", "down", "unoccupied",
}
ADMIN_WORDS = {"admin", "office", "model", "employee", "manager", "non revenue"}
TOTAL_WORDS = {
    "total", "totals", "subtotal", "grand total", "property total",
    "report total", "summary",
}


@dataclass
class MappingDecision:
    source_index: int
    source_header: str
    canonical_field: str
    confidence: float
    method: str
    reason: str


@dataclass
class ChargeDecision:
    source_text: str
    category: str
    confidence: float
    method: str
    reason: str


@dataclass
class SheetData:
    name: str
    rows: list[list[Any]]
    max_rows: int
    max_columns: int


@dataclass
class RunContext:
    input_path: Path
    run_dir: Path
    model: str
    save_intermediate: bool
    debug: bool
    logger: logging.Logger
    warnings: list[str] = field(default_factory=list)

    def artifact(self, filename: str, payload: Any) -> None:
        if not self.save_intermediate and filename != "final_output.json":
            return
        path = self.run_dir / filename
        with path.open("w", encoding="utf-8") as handle:
            json.dump(json_safe(payload), handle, indent=2, ensure_ascii=False)
            handle.write("\n")
        self.logger.debug("Saved intermediate artifact %s", filename)

    def warn(self, message: str) -> None:
        self.warnings.append(message)
        self.logger.warning(message)


class OpenRouterClient:
    def __init__(self, model: str, logger: logging.Logger) -> None:
        self.api_key = os.getenv("OPENROUTER_API_KEY", "").strip()
        self.model = model
        self.logger = logger
        self.calls: list[dict[str, Any]] = []

    @property
    def available(self) -> bool:
        return bool(self.api_key)

    def json_call(
        self,
        task: str,
        system_prompt: str,
        user_payload: dict[str, Any],
    ) -> Optional[Any]:
        if not self.available:
            self.logger.info(
                "OpenRouter skipped for %s: OPENROUTER_API_KEY is not set", task
            )
            self.calls.append({"task": task, "status": "skipped", "reason": "no_api_key"})
            return None

        body = {
            "model": self.model,
            "temperature": 0,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": json.dumps(json_safe(user_payload), ensure_ascii=False),
                },
            ],
        }
        request = urllib.request.Request(
            OPENROUTER_URL,
            data=json.dumps(body).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": os.getenv(
                    "OPENROUTER_SITE_URL", "https://localhost/rent-roll-standardizer"
                ),
                "X-Title": "Rent Roll Standardizer v0",
            },
            method="POST",
        )
        started = time.monotonic()
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                response_payload = json.loads(response.read().decode("utf-8"))
            content = response_payload["choices"][0]["message"]["content"]
            parsed = parse_json_object(content)
            elapsed_ms = round((time.monotonic() - started) * 1000)
            usage = response_payload.get("usage", {})
            self.calls.append(
                {
                    "task": task,
                    "status": "ok",
                    "model": response_payload.get("model", self.model),
                    "elapsed_ms": elapsed_ms,
                    "usage": usage,
                }
            )
            self.logger.info(
                "OpenRouter completed %s using %s in %d ms",
                task,
                response_payload.get("model", self.model),
                elapsed_ms,
            )
            return parsed
        except (urllib.error.URLError, TimeoutError, KeyError, ValueError, json.JSONDecodeError) as exc:
            self.calls.append(
                {"task": task, "status": "error", "error": str(exc), "model": self.model}
            )
            self.logger.warning("OpenRouter %s failed; deterministic fallback used: %s", task, exc)
            return None


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Standardize one CSV/XLSX/XLS rent roll into explainable JSON.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input", required=True, type=Path, help="Source .csv, .xlsx, or .xls")
    parser.add_argument("--output-dir", type=Path, default=Path("./output"))
    parser.add_argument(
        "--model",
        default=os.getenv("OPENROUTER_MODEL", DEFAULT_MODEL),
        help="OpenRouter model ID; OPENROUTER_MODEL can also set it",
    )
    parser.add_argument("--sheet", help="Manual worksheet name override")
    parser.add_argument("--debug", action="store_true", help="Include raw cell values in artifacts")
    parser.add_argument("--verbose", action="store_true", help="Print DEBUG logs to the console")
    parser.add_argument(
        "--save-intermediate",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Write numbered stage artifacts",
    )
    return parser.parse_args(argv)


def load_dotenv() -> Optional[Path]:
    """Load simple KEY=VALUE settings from a local .env without dependencies."""
    candidates = [Path.cwd() / ".env", Path(__file__).resolve().parent / ".env"]
    seen: set[Path] = set()
    for path in candidates:
        path = path.resolve()
        if path in seen or not path.is_file():
            continue
        seen.add(path)
        for line_number, raw_line in enumerate(
            path.read_text(encoding="utf-8-sig").splitlines(), start=1
        ):
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[7:].strip()
            if "=" not in line:
                print(
                    f"WARNING: Ignored malformed .env line {line_number}",
                    file=sys.stderr,
                )
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
                print(
                    f"WARNING: Ignored invalid .env key on line {line_number}",
                    file=sys.stderr,
                )
                continue
            if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
                value = value[1:-1]
            os.environ.setdefault(key, value)
        return path
    return None


def setup_run(args: argparse.Namespace) -> RunContext:
    input_path = args.input.expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
        raise ValueError("Supported input types are .csv, .xlsx, and .xls")

    output_root = args.output_dir.expanduser().resolve()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = slugify(input_path.stem) or "rent_roll"
    run_dir = output_root / f"{stem}_{timestamp}"
    counter = 2
    while run_dir.exists():
        run_dir = output_root / f"{stem}_{timestamp}_{counter}"
        counter += 1
    run_dir.mkdir(parents=True)

    logger = logging.getLogger(f"rent_roll_standardizer.{timestamp}.{counter}")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(run_dir / "debug.log", encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG if args.verbose else logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(console_handler)

    return RunContext(
        input_path=input_path,
        run_dir=run_dir,
        model=args.model,
        save_intermediate=args.save_intermediate,
        debug=args.debug,
        logger=logger,
    )


def load_workbook(path: Path) -> tuple[list[SheetData], dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        rows, encoding, delimiter = load_csv(path)
        sheets = [
            SheetData(
                name=path.stem,
                rows=rows,
                max_rows=len(rows),
                max_columns=max((len(row) for row in rows), default=0),
            )
        ]
        metadata = {
            "file_type": "csv",
            "encoding": encoding,
            "delimiter": delimiter,
            "sheet_count": 1,
        }
        return sheets, metadata
    if suffix == ".xlsx":
        try:
            import openpyxl  # type: ignore
        except ImportError as exc:
            raise RuntimeError("Reading .xlsx requires: pip install openpyxl") from exc
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets = []
        for worksheet in workbook.worksheets:
            rows = [
                [normalize_loaded_cell(value) for value in row]
                for row in worksheet.iter_rows(values_only=True)
            ]
            rows = trim_matrix(rows)
            sheets.append(
                SheetData(
                    name=worksheet.title,
                    rows=rows,
                    max_rows=len(rows),
                    max_columns=max((len(row) for row in rows), default=0),
                )
            )
        return sheets, {
            "file_type": "xlsx",
            "sheet_count": len(sheets),
            "workbook_properties": {
                "title": workbook.properties.title,
                "creator": workbook.properties.creator,
            },
        }
    try:
        import xlrd  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Reading legacy .xls requires: pip install xlrd") from exc
    workbook = xlrd.open_workbook(path, on_demand=True)
    sheets = []
    for worksheet in workbook.sheets():
        rows = []
        for row_index in range(worksheet.nrows):
            values = []
            for col_index in range(worksheet.ncols):
                cell = worksheet.cell(row_index, col_index)
                value: Any = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = datetime(*xlrd.xldate_as_tuple(value, workbook.datemode))
                values.append(normalize_loaded_cell(value))
            rows.append(values)
        rows = trim_matrix(rows)
        sheets.append(
            SheetData(
                name=worksheet.name,
                rows=rows,
                max_rows=len(rows),
                max_columns=max((len(row) for row in rows), default=0),
            )
        )
    return sheets, {"file_type": "xls", "sheet_count": len(sheets)}


def load_csv(path: Path) -> tuple[list[list[Any]], str, str]:
    raw = path.read_bytes()
    encoding = "utf-8-sig"
    try:
        text = raw.decode(encoding)
    except UnicodeDecodeError:
        encoding = "cp1252"
        text = raw.decode(encoding)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ","
    rows = [
        [normalize_loaded_cell(cell) for cell in row]
        for row in csv.reader(text.splitlines(), delimiter=delimiter)
    ]
    return trim_matrix(rows), encoding, delimiter


def trim_matrix(rows: list[list[Any]]) -> list[list[Any]]:
    while rows and not any(not is_blank(value) for value in rows[-1]):
        rows.pop()
    max_col = 0
    for row in rows:
        for index, value in enumerate(row):
            if not is_blank(value):
                max_col = max(max_col, index + 1)
    return [row[:max_col] + [None] * max(0, max_col - len(row)) for row in rows]


def select_sheet(
    sheets: list[SheetData], override: Optional[str]
) -> tuple[SheetData, list[dict[str, Any]], str]:
    if not sheets:
        raise ValueError("The input contains no readable sheets")
    scored = []
    for sheet in sheets:
        header_candidates = score_header_rows(sheet.rows)
        best_header = header_candidates[0] if header_candidates else {"score": 0.0}
        nonempty_rows = sum(1 for row in sheet.rows if any(not is_blank(v) for v in row))
        name_bonus = 0.15 if re.search(r"rent|roll|unit|tenant", sheet.name, re.I) else 0
        score = min(
            1.0,
            best_header.get("score", 0.0) * 0.75
            + min(nonempty_rows / 50, 1) * 0.10
            + name_bonus,
        )
        scored.append(
            {
                "sheet_name": sheet.name,
                "score": round(score, 3),
                "rows": sheet.max_rows,
                "columns": sheet.max_columns,
                "best_header_row": best_header.get("row_number"),
                "header_score": best_header.get("score", 0.0),
                "selected": False,
            }
        )
    if override:
        matches = [sheet for sheet in sheets if sheet.name.casefold() == override.casefold()]
        if not matches:
            raise ValueError(
                f"Sheet {override!r} not found. Available: {', '.join(s.name for s in sheets)}"
            )
        selected = matches[0]
        reason = "manual --sheet override"
    else:
        best = max(scored, key=lambda item: item["score"])
        selected = next(sheet for sheet in sheets if sheet.name == best["sheet_name"])
        reason = "highest rent-roll likelihood score"
    for item in scored:
        item["selected"] = item["sheet_name"] == selected.name
    return selected, scored, reason


def score_header_rows(rows: list[list[Any]], limit: int = 60) -> list[dict[str, Any]]:
    candidates = []
    alias_terms = {
        normalize_text(alias)
        for aliases in FIELD_ALIASES.values()
        for alias in aliases
    }
    for index, row in enumerate(rows[:limit]):
        nonempty = [value for value in row if not is_blank(value)]
        if len(nonempty) < 2:
            continue
        normalized = [normalize_text(value) for value in nonempty]
        text_like = sum(not looks_numeric(value) for value in nonempty) / len(nonempty)
        alias_hits = sum(
            1
            for value in normalized
            if value in alias_terms or any(term in value for term in alias_terms if len(term) >= 5)
        )
        unique_ratio = len(set(normalized)) / len(normalized)
        following_rows = rows[index + 1:index + 5]
        data_density = 0.0
        if following_rows:
            data_density = sum(
                sum(not is_blank(v) for v in following) / max(len(row), 1)
                for following in following_rows
            ) / len(following_rows)
        score = min(
            1.0,
            0.10
            + min(alias_hits / 4, 1) * 0.55
            + text_like * 0.15
            + unique_ratio * 0.08
            + min(data_density, 1) * 0.12,
        )
        candidates.append(
            {
                "row_number": index + 1,
                "score": round(score, 3),
                "alias_hits": alias_hits,
                "nonempty_cells": len(nonempty),
                "values": [display_value(value) for value in row],
            }
        )
    return sorted(candidates, key=lambda item: (-item["score"], item["row_number"]))


def detect_header(
    rows: list[list[Any]],
) -> tuple[int, list[str], list[list[Any]], float, list[dict[str, Any]]]:
    candidates = score_header_rows(rows)
    if not candidates:
        raise ValueError("Could not find a plausible header row")
    best = candidates[0]
    base_header_index = int(best["row_number"]) - 1

    start = base_header_index
    header_nonempty = sum(not is_blank(value) for value in rows[base_header_index])
    for candidate_index in range(max(0, base_header_index - 2), base_header_index):
        candidate_row = rows[candidate_index]
        nonempty = sum(not is_blank(value) for value in candidate_row)
        # Multi-row headers normally occupy a meaningful portion of the same
        # columns. This avoids folding metadata pairs such as
        # "Report Date | 2026-06-30" into the actual column names.
        if (
            nonempty >= max(2, round(header_nonempty * 0.4))
            and not row_is_title(candidate_row)
        ):
            start = candidate_index
            break

    end = base_header_index
    for candidate_index in range(
        base_header_index + 1, min(len(rows), base_header_index + 3)
    ):
        candidate_row = rows[candidate_index]
        nonempty_values = [value for value in candidate_row if not is_blank(value)]
        if len(nonempty_values) < 2 or row_is_title(candidate_row):
            break
        text_ratio = sum(not looks_numeric(value) for value in nonempty_values) / len(nonempty_values)
        header_terms = sum(
            any(
                normalize_text(alias) in normalize_text(value)
                for aliases in FIELD_ALIASES.values()
                for alias in aliases
            )
            for value in nonempty_values
        )
        if text_ratio >= 0.8 and header_terms >= 1:
            end = candidate_index
        else:
            break

    raw_header_rows = rows[start:end + 1]
    width = max((len(row) for row in raw_header_rows), default=0)
    headers = []
    for column in range(width):
        parts = []
        for row in raw_header_rows:
            value = row[column] if column < len(row) else None
            normalized = clean_header_part(value)
            if normalized and (not parts or normalized.casefold() != parts[-1].casefold()):
                parts.append(normalized)
        headers.append(" | ".join(parts) if parts else f"unnamed_{column + 1}")
    return end, headers, raw_header_rows, float(best["score"]), candidates[:8]


def deterministic_column_mapping(headers: list[str]) -> list[MappingDecision]:
    proposals: list[tuple[int, str, str, float, str]] = []
    for index, header in enumerate(headers):
        normalized = normalize_text(header.replace("|", " "))
        best_field = "unknown"
        best_score = 0.0
        best_alias = ""
        forced_field = distinctive_header_field(normalized)
        for field_name, aliases in FIELD_ALIASES.items():
            for alias in aliases:
                alias_normalized = normalize_text(alias)
                if normalized == alias_normalized:
                    score = 1.0
                elif re.search(rf"\b{re.escape(alias_normalized)}\b", normalized):
                    score = 0.88 if len(alias_normalized) >= 4 else 0.72
                elif token_similarity(normalized, alias_normalized) >= 0.75:
                    score = 0.70
                else:
                    continue
                if score > best_score or (
                    score == best_score
                    and len(alias_normalized.split()) > len(normalize_text(best_alias).split())
                ):
                    best_field, best_score, best_alias = field_name, score, alias
        if forced_field:
            best_field = forced_field
            best_score = max(best_score, 0.94)
            best_alias = f"distinctive header pattern for {forced_field}"
        proposals.append((index, header, best_field, best_score, best_alias))

    winners: dict[str, int] = {}
    for index, header, field_name, score, _ in proposals:
        if score < 0.70 or field_name in {
            "unknown", "charge_code", "charge_description", "charge_amount"
        }:
            continue
        preference = mapping_preference(field_name, header)
        current_index = winners.get(field_name)
        if current_index is None:
            winners[field_name] = index
            continue
        current = proposals[current_index]
        current_rank = (mapping_preference(field_name, current[1]), current[3])
        if (preference, score) > current_rank:
            winners[field_name] = index

    decisions = []
    for index, header, best_field, best_score, best_alias in proposals:
        if best_score < 0.70:
            decisions.append(
                MappingDecision(
                    index, header, "unknown", 0.2, "unresolved",
                    "No confident deterministic alias match",
                )
            )
        elif (
            best_field not in {"charge_code", "charge_description", "charge_amount"}
            and winners.get(best_field) != index
        ):
            decisions.append(
                MappingDecision(
                    index, header, "unknown", 0.3, "deterministic",
                    f"Another column was a better candidate for {best_field}",
                )
            )
        else:
            decisions.append(
                MappingDecision(
                    index,
                    header,
                    best_field,
                    round(best_score, 2),
                    "deterministic",
                    f"Matched alias {best_alias!r}",
                )
            )
    return decisions


def ai_map_columns(
    decisions: list[MappingDecision],
    sample_rows: list[list[Any]],
    client: OpenRouterClient,
) -> list[MappingDecision]:
    unresolved = [decision for decision in decisions if decision.canonical_field == "unknown"]
    if not unresolved:
        return decisions
    mapped_fields = {
        decision.canonical_field
        for decision in decisions
        if decision.canonical_field not in {"unknown", "ignore"}
    }
    payload = {
        "columns": [
            {
                "index": decision.source_index,
                "header": decision.source_header,
                "sample_values": [
                    display_value(row[decision.source_index])
                    for row in sample_rows[:8]
                    if decision.source_index < len(row)
                    and not is_blank(row[decision.source_index])
                ][:5],
            }
            for decision in unresolved
        ],
        "already_mapped_fields": sorted(mapped_fields),
        "allowed_fields": sorted(CANONICAL_FIELDS),
    }
    result = client.json_call(
        "column_mapping",
        (
            "You map ambiguous rent-roll columns. Return JSON only as "
            '{"mappings":[{"index":0,"field":"tenant_name","confidence":0.82,'
            '"reason":"..."}]}. Use only allowed_fields. Use unknown when uncertain. '
            "Do not calculate or transform values."
        ),
        payload,
    )
    if not isinstance(result, dict) or not isinstance(result.get("mappings"), list):
        return decisions
    by_index = {decision.source_index: decision for decision in decisions}
    claimed = set(mapped_fields)
    for item in result["mappings"]:
        try:
            index = int(item["index"])
            field_name = str(item["field"])
            confidence = float(item.get("confidence", 0.5))
        except (KeyError, TypeError, ValueError):
            continue
        if index not in by_index or field_name not in CANONICAL_FIELDS:
            continue
        if field_name in claimed and field_name not in {
            "unknown", "ignore", "charge_code", "charge_description", "charge_amount"
        }:
            continue
        confidence = min(max(confidence, 0.0), 0.84)
        if confidence < 0.55:
            field_name = "unknown"
        if field_name not in {"unknown", "ignore"}:
            claimed.add(field_name)
        original = by_index[index]
        by_index[index] = MappingDecision(
            index,
            original.source_header,
            field_name,
            round(confidence, 2),
            "openrouter",
            str(item.get("reason", "AI semantic mapping")),
        )
    return [by_index[index] for index in sorted(by_index)]


def detect_structure(
    rows: list[list[Any]],
    header_index: int,
    mappings: list[MappingDecision],
) -> dict[str, Any]:
    field_to_index = mapping_indexes(mappings)
    data_rows = rows[header_index + 1:]
    nonempty_data_rows = [
        row for row in data_rows if any(not is_blank(value) for value in row)
    ]
    unit_index = field_to_index.get("unit_number")
    explicit_unit_rows = 0
    continuation_rows = 0
    total_rows = 0
    if unit_index is not None:
        seen_unit = False
        for row in nonempty_data_rows:
            if is_total_row(row):
                total_rows += 1
                continue
            unit_value = cell(row, unit_index)
            if not is_blank(unit_value):
                explicit_unit_rows += 1
                seen_unit = True
            elif seen_unit:
                continuation_rows += 1
    has_charge_triplet = "charge_amount" in field_to_index and (
        "charge_code" in field_to_index or "charge_description" in field_to_index
    )
    if unit_index is None:
        pattern = "unresolved_unit_column"
        confidence = 0.15
    elif continuation_rows > max(2, explicit_unit_rows * 0.15) and has_charge_triplet:
        pattern = "repeating_unit_blocks"
        confidence = 0.88
    else:
        pattern = "flat_unit_rows"
        confidence = 0.90 if explicit_unit_rows else 0.35
    return {
        "pattern": pattern,
        "confidence": confidence,
        "data_start_row": header_index + 2,
        "unit_column_index": unit_index,
        "explicit_unit_rows": explicit_unit_rows,
        "continuation_rows": continuation_rows,
        "total_or_footer_rows": total_rows,
        "charge_columns_detected": has_charge_triplet,
        "rows_considered": len(nonempty_data_rows),
    }


def group_units(
    rows: list[list[Any]],
    header_index: int,
    mappings: list[MappingDecision],
    context: RunContext,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    indexes = mapping_indexes(mappings)
    unit_index = indexes.get("unit_number")
    if unit_index is None:
        context.warn("Could not confidently map a unit-number column; no units were grouped")
        return [], []

    groups: dict[str, dict[str, Any]] = {}
    order: list[str] = []
    skipped: list[dict[str, Any]] = []
    current_unit: Optional[str] = None
    future_unit: Optional[str] = None
    active_section = "primary"
    for row_index, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        if not any(not is_blank(value) for value in row):
            continue
        if is_property_total_row(row, indexes):
            skipped.append(
                {"row_number": row_index, "reason": "property_total_footer", "preview": row_preview(row)}
            )
            break
        raw_unit = cell(row, unit_index)
        normalized_unit = standardize_unit_id(raw_unit)
        row_has_charge = (
            not is_blank(cell(row, indexes.get("charge_amount")))
            or not is_blank(cell(row, indexes.get("charge_code")))
            or not is_blank(cell(row, indexes.get("charge_description")))
        )
        if is_section_row(row) and not row_has_charge:
            section_text = normalize_text(" ".join(display_value(value) for value in row))
            active_section = (
                "future"
                if "future" in section_text or "applicant" in section_text
                else "primary"
            )
            skipped.append(
                {
                    "row_number": row_index,
                    "reason": f"{active_section}_section_heading",
                    "preview": row_preview(row),
                }
            )
            continue
        if active_section == "future":
            if normalized_unit:
                future_unit = normalized_unit
            if future_unit and future_unit in groups:
                groups[future_unit]["future_source_rows"].append(row_index)
            skipped.append(
                {
                    "row_number": row_index,
                    "reason": "future_resident_record",
                    "associated_unit": future_unit,
                    "preview": row_preview(row),
                }
            )
            continue
        if is_total_row(row):
            if current_unit and current_unit in groups:
                reported_total = money_decimal(cell(row, indexes.get("charge_amount")))
                if reported_total is not None:
                    groups[current_unit]["reported_charge_total"] = reported_total
            skipped.append({"row_number": row_index, "reason": "total_or_footer", "preview": row_preview(row)})
            continue
        if normalized_unit:
            current_unit = normalized_unit
            if current_unit not in groups:
                groups[current_unit] = {
                    "unit_number": current_unit,
                    "source_rows": [],
                    "raw_unit_values": [],
                    "reported_charge_total": None,
                    "future_source_rows": [],
                }
                order.append(current_unit)
            groups[current_unit]["raw_unit_values"].append(display_value(raw_unit))
        elif current_unit and row_has_charge:
            pass
        else:
            reason = "continuation_without_unit" if current_unit else "row_before_first_unit"
            skipped.append({"row_number": row_index, "reason": reason, "preview": row_preview(row)})
            continue
        groups[current_unit]["source_rows"].append(
            {"row_number": row_index, "values": row}
        )

    compact_groups = []
    for unit_number in order:
        group = groups[unit_number]
        compact_groups.append(
            {
                "unit_number": unit_number,
                "source_row_numbers": [item["row_number"] for item in group["source_rows"]],
                "row_count": len(group["source_rows"]),
                "raw_unit_values": unique_preserving_order(group["raw_unit_values"]),
                "reported_charge_total": money_string(group.get("reported_charge_total")),
                "future_source_rows": group["future_source_rows"],
                **(
                    {"rows": [item["values"] for item in group["source_rows"]]}
                    if context.debug else {}
                ),
            }
        )
    return [groups[unit_number] for unit_number in order], compact_groups + [
        {"skipped_rows": skipped}
    ]


def classify_charges(
    groups: list[dict[str, Any]],
    mappings: list[MappingDecision],
    client: OpenRouterClient,
    context: RunContext,
) -> tuple[dict[str, ChargeDecision], list[dict[str, Any]]]:
    indexes = mapping_indexes(mappings)
    code_index = indexes.get("charge_code")
    description_index = indexes.get("charge_description")
    amount_index = indexes.get("charge_amount")
    if amount_index is None or (code_index is None and description_index is None):
        return {}, []

    labels: list[str] = []
    label_samples: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for group in groups:
        for source_row in group["source_rows"]:
            row = source_row["values"]
            label = charge_label(row, code_index, description_index)
            if label and label not in labels:
                labels.append(label)
            if label and len(label_samples[label]) < 3:
                label_samples[label].append(
                    {
                        "unit": group["unit_number"],
                        "amount": display_value(cell(row, amount_index)),
                    }
                )

    decisions: dict[str, ChargeDecision] = {}
    unresolved = []
    for label in labels:
        category, confidence, reason = deterministic_charge_category(label)
        if category:
            decisions[label] = ChargeDecision(label, category, confidence, "deterministic", reason)
        else:
            unresolved.append(label)

    if unresolved:
        result = client.json_call(
            "charge_classification",
            (
                "Classify rent-roll charge labels. Return JSON only as "
                '{"classifications":[{"source_text":"rnt","category":"BASE_RENT",'
                '"confidence":0.9,"reason":"..."}]}. Allowed categories are BASE_RENT, '
                "CONCESSION, UTILITY, PARKING, PET_FEE, PEST_FEE, TRASH_FEE, "
                "STORAGE, ADMIN_FEE, OTHER_RENT, ONE_TIME_FEE, DEPOSIT, UNKNOWN. "
                "Do not calculate amounts."
            ),
            {
                "charges": [
                    {"source_text": label, "samples": label_samples[label]}
                    for label in unresolved
                ]
            },
        )
        if isinstance(result, dict) and isinstance(result.get("classifications"), list):
            for item in result["classifications"]:
                label = str(item.get("source_text", ""))
                category = str(item.get("category", "UNKNOWN"))
                if label not in unresolved or category not in (
                    set(CHARGE_ALIASES) | {"UNKNOWN"}
                ):
                    continue
                try:
                    confidence = min(max(float(item.get("confidence", 0.5)), 0), 0.84)
                except (TypeError, ValueError):
                    confidence = 0.5
                if confidence < 0.55:
                    category = "UNKNOWN"
                decisions[label] = ChargeDecision(
                    label, category, round(confidence, 2), "openrouter",
                    str(item.get("reason", "AI semantic classification")),
                )

    for label in unresolved:
        if label not in decisions:
            decisions[label] = ChargeDecision(
                label, "UNKNOWN", 0.2, "unresolved",
                "No deterministic match and AI was unavailable or uncertain",
            )
            context.warn(f"Ambiguous charge code found: {label!r}")

    artifact = [
        {
            **asdict(decision),
            "sample_occurrences": label_samples[decision.source_text],
        }
        for decision in decisions.values()
    ]
    return decisions, artifact


def normalize_units(
    groups: list[dict[str, Any]],
    mappings: list[MappingDecision],
    charge_decisions: dict[str, ChargeDecision],
    context: RunContext,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    indexes = mapping_indexes(mappings)
    units = []
    calculations = []
    for group in groups:
        flags: list[dict[str, Any]] = []
        rows = [source["values"] for source in group["source_rows"]]
        source_rows = [source["row_number"] for source in group["source_rows"]]

        raw_tenant_name = first_value(rows, indexes.get("tenant_name"))
        unit: dict[str, Any] = {
            "unit_number": group["unit_number"],
            "unit_type": first_value(rows, indexes.get("unit_type")),
            "unit_sqft": number_value(first_value(rows, indexes.get("unit_sqft"))),
            "tenant_name": normalize_tenant_name(raw_tenant_name),
            "status": normalize_status(
                first_value(rows, indexes.get("status")),
                raw_tenant_name,
            ),
            "market_rent": money_value(first_value(rows, indexes.get("market_rent"))),
            "effective_rent": None,
            "total_rent": None,
            "move_in_date": date_value(first_value(rows, indexes.get("move_in_date"))),
            "lease_start_date": date_value(first_value(rows, indexes.get("lease_start_date"))),
            "lease_end_date": date_value(first_value(rows, indexes.get("lease_end_date"))),
            "move_out_date": date_value(first_value(rows, indexes.get("move_out_date"))),
            "balance": money_value(first_value(rows, indexes.get("balance"))),
            "deposit": money_value(first_value(rows, indexes.get("deposit"))),
            "charges": [],
            "flags": flags,
            "confidence": {},
            "source": {
                "sheet_rows": source_rows,
                "raw_unit_values": unique_preserving_order(group["raw_unit_values"]),
                "reported_charge_total": money_string(
                    money_decimal(group.get("reported_charge_total"))
                ),
                "future_record_rows": group.get("future_source_rows", []),
            },
        }
        if unit["status"] == "occupied" and not unit["lease_end_date"]:
            flags.append(
                flag("MISSING_LEASE_END", "Occupied unit had no lease end date", "warning")
            )
        if unit["status"] == "admin_model":
            flags.append(
                flag("ADMIN_MODEL_UNIT", "Unit was identified as admin/model", "info")
            )
        if group.get("future_source_rows"):
            flags.append(
                flag(
                    "FUTURE_RESIDENT_RECORD",
                    "A future resident/applicant record exists in the source",
                    "info",
                )
            )

        charge_amount_index = indexes.get("charge_amount")
        charge_code_index = indexes.get("charge_code")
        charge_description_index = indexes.get("charge_description")
        for source in group["source_rows"]:
            row = source["values"]
            amount = money_decimal(cell(row, charge_amount_index))
            label = charge_label(row, charge_code_index, charge_description_index)
            if amount is None or not label:
                continue
            decision = charge_decisions.get(
                label,
                ChargeDecision(label, "UNKNOWN", 0.2, "unresolved", "No classification"),
            )
            unit["charges"].append(
                {
                    "source_text": label,
                    "category": decision.category,
                    "amount": money_string(amount),
                    "source_row": source["row_number"],
                    "confidence": decision.confidence,
                }
            )
            if decision.category == "UNKNOWN":
                flags.append(
                    flag(
                        "UNCERTAIN_CHARGE",
                        f"Charge {label!r} was not confidently classified",
                        "warning",
                        source["row_number"],
                    )
                )

        explicit_effective = money_decimal(
            first_value(rows, indexes.get("effective_rent"))
        )
        explicit_total = money_decimal(first_value(rows, indexes.get("total_rent")))
        reported_charge_total = money_decimal(group.get("reported_charge_total"))
        if explicit_total is None:
            explicit_total = reported_charge_total
        charges = unit["charges"]
        base_total = sum_money(
            money_decimal(charge["amount"])
            for charge in charges
            if charge["category"] == "BASE_RENT"
        )
        concession_total = sum_money(
            normalized_concession(money_decimal(charge["amount"]))
            for charge in charges
            if charge["category"] == "CONCESSION"
        )
        recurring_total = sum_money(
            money_decimal(charge["amount"])
            for charge in charges
            if charge["category"] in RECURRING_CHARGE_CATEGORIES
        )

        effective: Optional[Decimal]
        total: Optional[Decimal]
        effective_formula: str
        total_formula: str
        if base_total is not None:
            effective = (base_total + (concession_total or Decimal("0"))).quantize(MONEY_QUANTUM)
            effective_formula = "sum(BASE_RENT) + normalized sum(CONCESSION)"
        elif explicit_effective is not None:
            effective = explicit_effective
            effective_formula = "source effective_rent field"
        elif explicit_total is not None:
            effective = explicit_total
            effective_formula = "fallback to source total_rent field"
        else:
            effective = None
            effective_formula = "unavailable"

        if recurring_total is not None:
            total = (recurring_total + (concession_total or Decimal("0"))).quantize(MONEY_QUANTUM)
            total_formula = "sum(recurring charges) + normalized sum(CONCESSION)"
        elif explicit_total is not None:
            total = explicit_total
            total_formula = "source total_rent field"
        elif effective is not None:
            total = effective
            total_formula = "fallback to calculated effective_rent"
        else:
            total = None
            total_formula = "unavailable"

        unit["effective_rent"] = money_string(effective)
        unit["total_rent"] = money_string(total)
        unit["market_rent"] = money_string(money_decimal(unit["market_rent"]))
        unit["balance"] = money_string(money_decimal(unit["balance"]))
        unit["deposit"] = money_string(money_decimal(unit["deposit"]))

        mapping_confidences = [
            decision.confidence
            for decision in mappings
            if decision.canonical_field not in {"unknown", "ignore"}
        ]
        structure_confidence = 0.95 if group["unit_number"] else 0.2
        semantic_confidence = (
            round(sum(mapping_confidences) / len(mapping_confidences), 2)
            if mapping_confidences else 0.2
        )
        charge_confidence = (
            round(
                sum(charge["confidence"] for charge in charges) / len(charges), 2
            )
            if charges else None
        )
        components = [structure_confidence, semantic_confidence]
        if charge_confidence is not None:
            components.append(charge_confidence)
        overall = round(sum(components) / len(components), 2)
        unit["confidence"] = {
            "overall": overall,
            "structure": structure_confidence,
            "semantic_mapping": semantic_confidence,
            "charge_classification": charge_confidence,
        }
        if overall < 0.65:
            flags.append(
                flag("LOW_CONFIDENCE", "Unit normalization confidence is low", "warning")
            )

        calculations.append(
            {
                "unit_number": group["unit_number"],
                "source_rows": source_rows,
                "base_rent_total": money_string(base_total),
                "concession_total": money_string(concession_total),
                "recurring_charge_total_before_concession": money_string(recurring_total),
                "effective_rent": money_string(effective),
                "effective_rent_formula": effective_formula,
                "total_rent": money_string(total),
                "total_rent_formula": total_formula,
            }
        )
        context.logger.debug(
            "Calculated effective rent for unit %s = %s",
            group["unit_number"],
            unit["effective_rent"] or "unavailable",
        )
        units.append(unit)
    return units, calculations


def validate(
    units: list[dict[str, Any]],
    rows: list[list[Any]],
    header_index: int,
    mappings: list[MappingDecision],
) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []

    add_check(
        checks,
        "reasonable_unit_count",
        bool(units),
        f"Found {len(units)} units",
        "error" if not units else "info",
    )
    counts = Counter(unit["unit_number"] for unit in units)
    duplicates = sorted(unit for unit, count in counts.items() if count > 1)
    add_check(
        checks,
        "duplicate_unit_numbers",
        not duplicates,
        "No duplicate normalized unit numbers" if not duplicates else f"Duplicates: {duplicates}",
        "warning",
    )

    invalid_dates = []
    for unit in units:
        for field_name in (
            "move_in_date", "lease_start_date", "lease_end_date", "move_out_date"
        ):
            value = unit.get(field_name)
            if value and not valid_iso_date(value):
                invalid_dates.append({"unit": unit["unit_number"], "field": field_name, "value": value})
    add_check(
        checks,
        "valid_dates",
        not invalid_dates,
        "All normalized dates are valid" if not invalid_dates else f"{len(invalid_dates)} invalid dates",
        "warning",
        invalid_dates,
    )

    rent_anomalies = []
    for unit in units:
        effective = money_decimal(unit.get("effective_rent"))
        total = money_decimal(unit.get("total_rent"))
        if effective is not None and total is not None and effective > total + Decimal("0.01"):
            rent_anomalies.append(
                {
                    "unit": unit["unit_number"],
                    "effective_rent": money_string(effective),
                    "total_rent": money_string(total),
                }
            )
    add_check(
        checks,
        "effective_not_above_total",
        not rent_anomalies,
        "Effective rent does not exceed total rent"
        if not rent_anomalies else f"{len(rent_anomalies)} rent anomalies",
        "warning",
        rent_anomalies,
    )

    unit_total_mismatches = []
    for unit in units:
        reported = money_decimal(unit.get("source", {}).get("reported_charge_total"))
        computed = money_decimal(unit.get("total_rent"))
        if reported is not None and computed is not None and abs(reported - computed) > Decimal("0.01"):
            unit_total_mismatches.append(
                {
                    "unit": unit["unit_number"],
                    "reported": money_string(reported),
                    "computed": money_string(computed),
                    "difference": money_string(abs(reported - computed)),
                }
            )
    add_check(
        checks,
        "unit_charge_total_reconciliation",
        not unit_total_mismatches,
        "Per-unit source totals match deterministic totals"
        if not unit_total_mismatches
        else f"{len(unit_total_mismatches)} unit charge totals did not match",
        "warning",
        unit_total_mismatches,
    )

    vacant_tenants = [
        {"unit": unit["unit_number"], "tenant_name": unit["tenant_name"]}
        for unit in units
        if unit["status"] == "vacant" and not is_blank(unit.get("tenant_name"))
    ]
    add_check(
        checks,
        "vacant_units_without_tenants",
        not vacant_tenants,
        "Vacant units do not carry tenant names"
        if not vacant_tenants else f"{len(vacant_tenants)} vacant units carry tenant names",
        "warning",
        vacant_tenants,
    )

    admin_unflagged = [
        unit["unit_number"]
        for unit in units
        if any(word in normalize_text(unit.get("status")) for word in ADMIN_WORDS)
        and unit.get("status") != "admin_model"
    ]
    add_check(
        checks,
        "admin_model_units_flagged",
        not admin_unflagged,
        "Admin/model units are flagged" if not admin_unflagged else f"Unflagged: {admin_unflagged}",
        "warning",
    )

    source_totals = detect_source_total_amounts(
        rows[header_index + 1:], mapping_indexes(mappings)
    )
    computed_total = sum_money(money_decimal(unit.get("total_rent")) for unit in units)
    total_comparison = compare_source_totals(source_totals, computed_total)
    checks.append(total_comparison)

    failed = sum(1 for check in checks if not check["passed"])
    return {
        "summary": {
            "checks_run": len(checks),
            "passed": len(checks) - failed,
            "failed": failed,
            "status": "passed" if failed == 0 else "completed_with_warnings",
        },
        "checks": checks,
    }


def build_rollup(
    units: list[dict[str, Any]],
    rent_roll_date: Optional[str],
    property_name: Optional[str],
) -> dict[str, Any]:
    occupied = sum(unit["status"] == "occupied" for unit in units)
    vacant = sum(unit["status"] == "vacant" for unit in units)
    admin_model = sum(unit["status"] == "admin_model" for unit in units)
    denomin = occupied + vacant
    effective_values = [
        value
        for unit in units
        if (value := money_decimal(unit.get("effective_rent"))) is not None
        and unit["status"] == "occupied"
    ]
    average = (
        sum(effective_values, Decimal("0")) / len(effective_values)
        if effective_values else None
    )
    return {
        "property_name": property_name,
        "total_units": len(units),
        "occupied_count": occupied,
        "vacant_count": vacant,
        "admin_model_count": admin_model,
        "unknown_status_count": len(units) - occupied - vacant - admin_model,
        "average_effective_rent": money_string(average),
        "property_occupancy_rate": (
            round(occupied / denomin, 4) if denomin else None
        ),
        "rent_roll_date": rent_roll_date,
    }


def detect_rent_roll_date(
    rows: list[list[Any]], header_index: int, mappings: list[MappingDecision]
) -> Optional[str]:
    indexes = mapping_indexes(mappings)
    date_index = indexes.get("rent_roll_date")
    if date_index is not None:
        for row in rows[header_index + 1:]:
            parsed = date_value(cell(row, date_index))
            if parsed:
                return parsed
    date_patterns = re.compile(
        r"(?:as\s+of|rent\s+roll\s+date|report\s+date)\s*[:=]?\s*(.+)", re.I
    )
    for row in rows[:header_index + 1]:
        for index, value in enumerate(row):
            if not isinstance(value, str):
                continue
            match = date_patterns.search(value)
            if match:
                parsed = date_value(match.group(1).strip())
                if parsed:
                    return parsed
            if normalize_text(value) in {"as of", "rent roll date", "report date"}:
                parsed = date_value(cell(row, index + 1))
                if parsed:
                    return parsed
    return None


def detect_property_name(
    rows: list[list[Any]],
    header_index: int,
    mappings: list[MappingDecision],
    input_path: Path,
) -> Optional[str]:
    indexes = mapping_indexes(mappings)
    property_index = indexes.get("property_name")
    if property_index is not None:
        for row in rows[header_index + 1:]:
            value = first_value([row], property_index)
            if not is_blank(value):
                return str(value).strip()

    generic = re.compile(
        r"^(?:rent\s*roll|rent\s*roll\s+detail|as\s+of|month\s+year|"
        r"report\s+date|parameters?|page\s+\d+)",
        re.I,
    )
    for row in rows[:header_index]:
        nonempty = [display_value(value).strip() for value in row if not is_blank(value)]
        if len(nonempty) != 1:
            continue
        candidate = nonempty[0]
        if (
            len(candidate) >= 3
            and not generic.search(candidate)
            and date_value(candidate) is None
            and not re.fullmatch(r"[\d\s:/=-]+", candidate)
        ):
            return candidate
    fallback = re.sub(r"[_-]+", " ", input_path.stem).strip()
    return fallback.title() if fallback else None


STANDARDIZED_COLUMNS = [
    "Unit No",
    "Tenant Name",
    "Status",
    "Unit Type",
    "Unit Size (SF)",
    "Market Rent",
    "Effective Rent",
    "Total Rent",
    "Move In Date",
    "Lease Start Date",
    "Lease End Date",
    "Move Out Date",
    "Balance",
    "Deposit",
    "Confidence",
    "Flags",
]


def flattened_unit_values(unit: dict[str, Any]) -> list[Any]:
    return [
        unit.get("unit_number"),
        unit.get("tenant_name"),
        unit.get("status"),
        unit.get("unit_type"),
        unit.get("unit_sqft"),
        unit.get("market_rent"),
        unit.get("effective_rent"),
        unit.get("total_rent"),
        unit.get("move_in_date"),
        unit.get("lease_start_date"),
        unit.get("lease_end_date"),
        unit.get("move_out_date"),
        unit.get("balance"),
        unit.get("deposit"),
        unit.get("confidence", {}).get("overall"),
        "; ".join(
            f"{item.get('code', 'FLAG')}: {item.get('message', '')}".rstrip(": ")
            for item in unit.get("flags", [])
        ),
    ]


def write_standardized_csv(canonical: dict[str, Any], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(STANDARDIZED_COLUMNS)
        for unit in canonical.get("units", []):
            writer.writerow(
                ["" if value is None else value for value in flattened_unit_values(unit)]
            )


def write_standardized_workbook(canonical: dict[str, Any], path: Path) -> None:
    try:
        from openpyxl import Workbook  # type: ignore
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore
        from openpyxl.utils import get_column_letter  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Writing standardized_rent_roll.xlsx requires: pip install openpyxl"
        ) from exc

    workbook = Workbook()
    detail = workbook.active
    detail.title = "Standardized Rent Roll"
    summary = workbook.create_sheet("Summary")

    navy = "17365D"
    teal = "0F6B78"
    white = "FFFFFF"
    light_blue = "DCE6F1"
    light_gray = "F5F7FA"
    border_gray = "D9E1E8"
    warning_fill = "FFF2CC"
    thin_gray = Side(style="thin", color=border_gray)

    detail.sheet_view.showGridLines = False
    detail.freeze_panes = "A2"
    detail.append(STANDARDIZED_COLUMNS)
    header = detail[1]
    for cell in header:
        cell.font = Font(name="Arial", size=9, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=Side(style="medium", color=teal))
    detail.row_dimensions[1].height = 26

    currency_columns = {6, 7, 8, 13, 14}
    date_columns = {9, 10, 11, 12}
    for row_number, unit in enumerate(canonical.get("units", []), start=2):
        values = flattened_unit_values(unit)
        excel_values = []
        for column_number, value in enumerate(values, start=1):
            if value is None or value == "":
                excel_values.append(None)
            elif column_number in currency_columns:
                excel_values.append(float(Decimal(str(value))))
            elif column_number in date_columns:
                excel_values.append(date.fromisoformat(str(value)))
            else:
                excel_values.append(value)
        detail.append(excel_values)
        for cell in detail[row_number]:
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center")
        if row_number % 2 == 0:
            for cell in detail[row_number]:
                cell.fill = PatternFill("solid", fgColor=light_gray)
        for column_number in currency_columns:
            detail.cell(row_number, column_number).number_format = (
                '$#,##0.00;[Red]-$#,##0.00'
            )
        for column_number in date_columns:
            detail.cell(row_number, column_number).number_format = "yyyy-mm-dd"
        detail.cell(row_number, 15).number_format = "0%"
        detail.cell(row_number, 5).number_format = '#,##0'
        detail.cell(row_number, 16).alignment = Alignment(wrap_text=True, vertical="top")
        status_cell = detail.cell(row_number, 3)
        status_color = {
            "occupied": "E2F0D9",
            "vacant": warning_fill,
            "admin_model": light_blue,
            "unknown": "E7E6E6",
        }.get(str(status_cell.value), None)
        if status_color:
            status_cell.fill = PatternFill("solid", fgColor=status_color)
        if detail.cell(row_number, 16).value:
            detail.cell(row_number, 16).fill = PatternFill("solid", fgColor=warning_fill)
            detail.row_dimensions[row_number].height = 30

    last_row = max(detail.max_row, 1)
    detail.auto_filter.ref = f"A1:P{last_row}"
    detail.sheet_properties.pageSetUpPr.fitToPage = True
    detail.page_setup.fitToWidth = 1
    detail.page_setup.fitToHeight = 0
    detail.sheet_properties.outlinePr.summaryBelow = True

    width_limits = {
        1: 14, 2: 24, 3: 14, 4: 18, 5: 15, 6: 16, 7: 17, 8: 15,
        9: 15, 10: 17, 11: 15, 12: 16, 13: 14, 14: 14, 15: 13, 16: 42,
    }
    for column_number, width in width_limits.items():
        detail.column_dimensions[get_column_letter(column_number)].width = width

    summary.sheet_view.showGridLines = False
    summary.merge_cells("A1:B1")
    summary["A1"] = "Property Summary"
    summary["A1"].font = Font(name="Arial", bold=True, size=16, color=white)
    summary["A1"].fill = PatternFill("solid", fgColor=navy)
    summary["A1"].alignment = Alignment(horizontal="left", vertical="center")
    summary.row_dimensions[1].height = 30

    rollup = canonical.get("rollup", {})
    validation_summary = canonical.get("validation_summary", {})
    flagged_units = sum(bool(unit.get("flags")) for unit in canonical.get("units", []))
    summary_rows = [
        ("Property Name", rollup.get("property_name")),
        ("Rent Roll Date", rollup.get("rent_roll_date")),
        ("Total Units", rollup.get("total_units")),
        ("Occupied Units", rollup.get("occupied_count")),
        ("Vacant Units", rollup.get("vacant_count")),
        ("Average Effective Rent", rollup.get("average_effective_rent")),
        ("Property Occupancy Rate", rollup.get("property_occupancy_rate")),
        ("Number of Flagged Units", flagged_units),
        ("Number of Validation Warnings", validation_summary.get("warning_count", 0)),
    ]
    for row_number, (label, value) in enumerate(summary_rows, start=3):
        label_cell = summary.cell(row_number, 1, label)
        label_cell.font = Font(name="Arial", size=10, bold=True, color=navy)
        label_cell.fill = PatternFill("solid", fgColor=light_blue)
        label_cell.border = Border(bottom=thin_gray)
        value_cell = summary.cell(row_number, 2)
        if value is None or value == "":
            value_cell.value = None
        elif label == "Rent Roll Date":
            value_cell.value = date.fromisoformat(str(value))
            value_cell.number_format = "yyyy-mm-dd"
        elif label == "Average Effective Rent":
            value_cell.value = float(Decimal(str(value)))
            value_cell.number_format = '$#,##0.00;[Red]-$#,##0.00'
        elif label == "Property Occupancy Rate":
            value_cell.value = float(value)
            value_cell.number_format = "0.0%"
        else:
            value_cell.value = value
        value_cell.border = Border(bottom=thin_gray)
        value_cell.font = Font(name="Arial", size=10)
        value_cell.alignment = Alignment(vertical="center")
        summary.row_dimensions[row_number].height = 23

    summary.column_dimensions["A"].width = 32
    summary.column_dimensions["B"].width = 42
    summary.freeze_panes = "A3"
    summary["A14"] = "Source of truth"
    summary["A14"].font = Font(name="Arial", size=10, bold=True, color=navy)
    summary["B14"] = "final_output.json (values are not recalculated in Excel)"
    summary["B14"].font = Font(name="Arial", size=10, italic=True, color="666666")

    workbook.save(path)


def run(args: argparse.Namespace) -> Path:
    context = setup_run(args)
    context.logger.info("Rent Roll Standardizer v%s", VERSION)
    context.logger.info("File loaded: %s", context.input_path)
    context.logger.info("AI model configured: %s", context.model)
    client = OpenRouterClient(context.model, context.logger)

    sheets, metadata = load_workbook(context.input_path)
    input_summary = {
        "input_file": str(context.input_path),
        "file_size_bytes": context.input_path.stat().st_size,
        "extension": context.input_path.suffix.lower(),
        "loaded_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "parser_version": VERSION,
        "model_configured": context.model,
        "openrouter_enabled": client.available,
        **metadata,
        "sheets": [
            {"name": sheet.name, "rows": sheet.max_rows, "columns": sheet.max_columns}
            for sheet in sheets
        ],
    }
    context.artifact("01_input_summary.json", input_summary)

    selected, sheet_scores, selection_reason = select_sheet(sheets, args.sheet)
    context.logger.info(
        "Detected %d sheet%s, selected %s (%s)",
        len(sheets), "" if len(sheets) == 1 else "s", selected.name, selection_reason,
    )
    context.artifact(
        "02_sheet_detection.json",
        {
            "selected_sheet": selected.name,
            "selection_reason": selection_reason,
            "candidates": sheet_scores,
        },
    )

    header_index, headers, raw_header_rows, header_confidence, header_candidates = detect_header(
        selected.rows
    )
    context.logger.info(
        "Header row appears to be row %d (confidence %.0f%%)",
        header_index + 1, header_confidence * 100,
    )
    context.artifact(
        "03_header_detection.json",
        {
            "header_row": header_index + 1,
            "confidence": header_confidence,
            "combined_headers": headers,
            "raw_header_rows": [
                [display_value(value) for value in row] for row in raw_header_rows
            ],
            "top_candidates": header_candidates,
        },
    )

    deterministic_mappings = deterministic_column_mapping(headers)
    mappings = ai_map_columns(
        deterministic_mappings, selected.rows[header_index + 1:header_index + 9], client
    )
    for decision in mappings:
        if decision.canonical_field == "unknown":
            context.warn(
                f"Could not confidently map column {decision.source_header!r}"
            )
        else:
            context.logger.info(
                "Mapped %r -> %s (%.0f%%, %s)",
                decision.source_header,
                decision.canonical_field,
                decision.confidence * 100,
                decision.method,
            )

    structure = detect_structure(selected.rows, header_index, mappings)
    context.logger.info(
        "Detected structural pattern %s (confidence %.0f%%)",
        structure["pattern"], structure["confidence"] * 100,
    )
    context.artifact(
        "04_structure_detection.json",
        {
            **structure,
            "fallback_used": structure["confidence"] < 0.65,
            "uncertain_columns": [
                decision.source_header
                for decision in mappings
                if decision.canonical_field == "unknown"
            ],
        },
    )

    groups, grouping_artifact = group_units(
        selected.rows, header_index, mappings, context
    )
    context.logger.info("Grouped %d units", len(groups))
    context.artifact(
        "05_unit_grouping.json",
        {
            "unit_count": len(groups),
            "grouping_method": structure["pattern"],
            "groups": grouping_artifact,
        },
    )
    context.artifact(
        "06_column_mapping.json",
        {
            "canonical_fields": sorted(CANONICAL_FIELDS),
            "decisions": [asdict(decision) for decision in mappings],
            "openrouter_calls": [
                call for call in client.calls if call["task"] == "column_mapping"
            ],
        },
    )

    charge_decisions, charge_artifact = classify_charges(
        groups, mappings, client, context
    )
    context.logger.info(
        "Categorized %d distinct charge label%s",
        len(charge_decisions), "" if len(charge_decisions) == 1 else "s",
    )
    context.artifact(
        "07_charge_classification.json",
        {
            "taxonomy": sorted(set(CHARGE_ALIASES) | {"UNKNOWN"}),
            "decisions": charge_artifact,
            "openrouter_calls": [
                call for call in client.calls if call["task"] == "charge_classification"
            ],
        },
    )

    units, calculations = normalize_units(
        groups, mappings, charge_decisions, context
    )
    context.logger.info("Calculated totals for %d units using deterministic rules", len(units))
    validation = validate(units, selected.rows, header_index, mappings)
    validation["summary"]["warning_count"] = sum(
        not check["passed"] and check.get("severity") == "warning"
        for check in validation["checks"]
    )
    for check in validation["checks"]:
        log_method = context.logger.info if check["passed"] else context.logger.warning
        log_method("Validation %s: %s", check["name"], check["message"])
    context.logger.info(
        "Validations run: %d checks, %d passed",
        validation["summary"]["checks_run"], validation["summary"]["passed"],
    )
    context.artifact(
        "08_validation_report.json",
        {
            **validation,
            "calculated_totals": calculations,
            "pipeline_warnings": context.warnings,
        },
    )

    rent_roll_date = detect_rent_roll_date(selected.rows, header_index, mappings)
    property_name = detect_property_name(
        selected.rows, header_index, mappings, context.input_path
    )
    output = {
        "schema_version": "rent_roll_standardizer.v0",
        "source": {
            "file_name": context.input_path.name,
            "selected_sheet": selected.name,
            "header_row": header_index + 1,
        },
        "processing": {
            "parser_version": VERSION,
            "deterministic_structure": True,
            "ai_scope": ["column_mapping", "charge_classification"],
            "openrouter_model": context.model if client.available else None,
            "openrouter_calls": client.calls,
            "warnings": context.warnings,
        },
        "rollup": build_rollup(units, rent_roll_date, property_name),
        "units": units,
        "validation_summary": validation["summary"],
    }
    context.artifact("final_output.json", output)
    workbook_path = context.run_dir / "standardized_rent_roll.xlsx"
    csv_path = context.run_dir / "standardized_rent_roll.csv"
    write_standardized_workbook(output, workbook_path)
    write_standardized_csv(output, csv_path)
    context.logger.info("Final output written: %s", context.run_dir / "final_output.json")
    context.logger.info("Standardized workbook written: %s", workbook_path)
    context.logger.info("Standardized CSV written: %s", csv_path)
    print(f"\nRun complete: {context.run_dir}")
    return context.run_dir


def mapping_indexes(mappings: list[MappingDecision]) -> dict[str, int]:
    return {
        decision.canonical_field: decision.source_index
        for decision in mappings
        if decision.canonical_field not in {"unknown", "ignore"}
    }


def mapping_preference(field_name: str, header: str) -> int:
    normalized = normalize_text(header.replace("|", " "))
    if field_name == "tenant_name":
        if "name" in normalized:
            return 3
        if "tenant" in normalized:
            return 2
        return 1
    if field_name == "deposit":
        if "resident" in normalized and "deposit" in normalized:
            return 3
        if normalized == "deposit" or "security deposit" in normalized:
            return 2
        return 1
    return 1


def distinctive_header_field(normalized_header: str) -> Optional[str]:
    if any(term in normalized_header for term in ("sq ft", "sqft", "square feet")):
        return "unit_sqft"
    if normalized_header.endswith("deposit"):
        return "deposit"
    if normalized_header.endswith("expiration") or normalized_header.endswith("expiry"):
        return "lease_end_date"
    return None


def deterministic_charge_category(label: str) -> tuple[Optional[str], float, str]:
    normalized = normalize_text(label)
    best: tuple[Optional[str], float, str] = (None, 0.0, "")
    for category, aliases in CHARGE_ALIASES.items():
        for alias in aliases:
            alias_normalized = normalize_text(alias)
            if normalized == alias_normalized:
                score = 1.0
            elif (
                len(alias_normalized) >= 4
                and re.search(rf"\b{re.escape(alias_normalized)}\b", normalized)
            ):
                score = 0.88
            else:
                continue
            if score > best[1]:
                best = (category, score, f"Matched charge alias {alias!r}")
    return best


def charge_label(
    row: list[Any], code_index: Optional[int], description_index: Optional[int]
) -> str:
    code = display_value(cell(row, code_index)).strip()
    description = display_value(cell(row, description_index)).strip()
    if code and description and normalize_text(code) != normalize_text(description):
        return f"{code} | {description}"
    return code or description


def normalize_status(value: Any, tenant_name: Any) -> str:
    normalized = normalize_text(value)
    tenant_normalized = normalize_text(tenant_name)
    if any(word in normalized or word in tenant_normalized for word in ADMIN_WORDS):
        return "admin_model"
    if any(word in normalized for word in VACANT_WORDS):
        return "vacant"
    if any(word in normalized for word in OCCUPIED_WORDS):
        return "occupied"
    if tenant_normalized in {"vacant", "vac", "available"}:
        return "vacant"
    if tenant_normalized:
        return "occupied"
    return "unknown"


def normalize_tenant_name(value: Any) -> Optional[str]:
    if is_blank(value):
        return None
    text = str(value).strip()
    if normalize_text(text) in VACANT_WORDS | {"n a", "none", "no tenant"}:
        return None
    return text


def standardize_unit_id(value: Any) -> Optional[str]:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    text = re.sub(r"\s*[-–—/]\s*", "-", text)
    text = re.sub(r"\s+", " ", text)
    if normalize_text(text) in TOTAL_WORDS:
        return None
    return text.upper()


def date_value(value: Any) -> Optional[str]:
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)) and 1 <= float(value) <= 100000:
        # Excel's 1900 date system, including its historical leap-year offset.
        from datetime import timedelta
        base = datetime(1899, 12, 30)
        try:
            return (base + timedelta(days=float(value))).date().isoformat()
        except (OverflowError, ValueError):
            return None
    text = str(value).strip()
    text = re.sub(r"^(?:as\s+of\s*:?\s*)", "", text, flags=re.I)
    formats = [
        "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y",
        "%m-%d-%y", "%d/%m/%Y", "%b %d, %Y", "%B %d, %Y", "%d-%b-%Y",
        "%b %Y", "%B %Y",
    ]
    for format_string in formats:
        try:
            return datetime.strptime(text, format_string).date().isoformat()
        except ValueError:
            continue
    return None


def valid_iso_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
        return True
    except (TypeError, ValueError):
        return False


def money_decimal(value: Any) -> Optional[Decimal]:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value.quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    text = str(value).strip()
    if not text or normalize_text(text) in {"n a", "na", "none", "-", "--"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[$£€₹,\s]", "", text.strip("()"))
    text = re.sub(r"[^0-9.+-]", "", text)
    try:
        amount = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    if negative:
        amount = -amount
    return amount.quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)


def money_value(value: Any) -> Optional[str]:
    return money_string(money_decimal(value))


def money_string(value: Optional[Decimal]) -> Optional[str]:
    if value is None:
        return None
    return format(value.quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP), "f")


def normalized_concession(value: Optional[Decimal]) -> Optional[Decimal]:
    if value is None:
        return None
    return -abs(value)


def sum_money(values: Iterable[Optional[Decimal]]) -> Optional[Decimal]:
    filtered = [value for value in values if value is not None]
    if not filtered:
        return None
    return sum(filtered, Decimal("0")).quantize(MONEY_QUANTUM)


def number_value(value: Any) -> Optional[float | int]:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else float(value)
    text = re.sub(r"[,\s]", "", str(value))
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(match.group())
    return int(number) if number.is_integer() else number


def first_value(rows: list[list[Any]], index: Optional[int]) -> Any:
    if index is None:
        return None
    for row in rows:
        value = cell(row, index)
        if not is_blank(value):
            return normalize_loaded_cell(value)
    return None


def cell(row: list[Any], index: Optional[int]) -> Any:
    if index is None or index < 0 or index >= len(row):
        return None
    return row[index]


def is_total_row(row: list[Any]) -> bool:
    populated = [normalize_text(value) for value in row if not is_blank(value)]
    if not populated:
        return False
    first_words = " ".join(populated[:2])
    return any(
        first_words == word or first_words.startswith(f"{word} ")
        for word in TOTAL_WORDS
    )


def is_property_total_row(row: list[Any], indexes: dict[str, int]) -> bool:
    if not is_blank(cell(row, indexes.get("unit_number"))):
        return False
    normalized_cells = [normalize_text(value) for value in row if not is_blank(value)]
    has_total_label = any(value in {"total", "totals", "grand total"} for value in normalized_cells)
    numeric_cells = sum(
        money_decimal(value) is not None for value in row if not is_blank(value)
    )
    return has_total_label and numeric_cells >= 3


def is_section_row(row: list[Any]) -> bool:
    populated = [normalize_text(value) for value in row if not is_blank(value)]
    if len(populated) > 2:
        return False
    text = " ".join(populated)
    return bool(
        re.search(
            r"\b(current|notice|vacant|future|former|applicant|resident)s?\b",
            text,
        )
        and ("resident" in text or "/" in display_value(next(v for v in row if not is_blank(v))))
    )


def detect_source_total_amounts(
    rows: list[list[Any]], indexes: Optional[dict[str, int]] = None
) -> list[dict[str, Any]]:
    totals = []
    for index, row in enumerate(rows, start=1):
        if indexes is not None:
            if not is_property_total_row(row, indexes):
                continue
        elif not is_total_row(row):
            continue
        amounts = [
            money_string(amount)
            for value in row
            if (amount := money_decimal(value)) is not None
        ]
        if amounts:
            totals.append({"relative_row": index, "amounts": amounts})
    return totals


def compare_source_totals(
    source_totals: list[dict[str, Any]], computed_total: Optional[Decimal]
) -> dict[str, Any]:
    if not source_totals:
        return {
            "name": "source_total_reconciliation",
            "passed": True,
            "severity": "info",
            "message": "No source total row was available for reconciliation",
            "details": {"computed_total_rent": money_string(computed_total)},
        }
    if computed_total is None:
        return {
            "name": "source_total_reconciliation",
            "passed": False,
            "severity": "warning",
            "message": "Source total rows exist but computed total rent is unavailable",
            "details": {"source_totals": source_totals},
        }
    candidates = [
        money_decimal(amount)
        for total in source_totals
        for amount in total["amounts"]
    ]
    differences = [
        abs(candidate - computed_total)
        for candidate in candidates if candidate is not None
    ]
    tolerance = max(Decimal("1.00"), abs(computed_total) * Decimal("0.0001"))
    matched = bool(differences) and min(differences) <= tolerance
    return {
        "name": "source_total_reconciliation",
        "passed": matched,
        "severity": "warning",
        "message": (
            "A source total matched computed total rent within tolerance"
            if matched else "Charge total row did not match computed total"
        ),
        "details": {
            "computed_total_rent": money_string(computed_total),
            "tolerance": money_string(tolerance),
            "source_totals": source_totals,
            "closest_difference": money_string(min(differences)) if differences else None,
        },
    }


def add_check(
    checks: list[dict[str, Any]],
    name: str,
    passed: bool,
    message: str,
    severity: str,
    details: Any = None,
) -> None:
    checks.append(
        {
            "name": name,
            "passed": passed,
            "severity": severity,
            "message": message,
            "details": details,
        }
    )


def flag(
    code: str,
    message: str,
    severity: str = "warning",
    source_row: Optional[int] = None,
) -> dict[str, Any]:
    result = {"code": code, "message": message, "severity": severity}
    if source_row is not None:
        result["source_row"] = source_row
    return result


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().casefold()
    text = re.sub(r"[_./\\|#-]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_header_part(value: Any) -> str:
    if is_blank(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def token_similarity(left: str, right: str) -> float:
    left_tokens = set(left.split())
    right_tokens = set(right.split())
    if not left_tokens or not right_tokens:
        return 0.0
    return len(left_tokens & right_tokens) / len(left_tokens | right_tokens)


def looks_numeric(value: Any) -> bool:
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return True
    return bool(re.fullmatch(r"[$£€₹(]?\s*[-+]?\d[\d,.]*\s*\)?%?", str(value).strip()))


def row_is_title(row: list[Any]) -> bool:
    nonempty = [value for value in row if not is_blank(value)]
    return len(nonempty) == 1


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def normalize_loaded_cell(value: Any) -> Any:
    if isinstance(value, str):
        return re.sub(r"\s+", " ", value.strip()) or None
    return value


def display_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def row_preview(row: list[Any], limit: int = 6) -> list[str]:
    return [display_value(value) for value in row[:limit]]


def unique_preserving_order(values: Iterable[Any]) -> list[Any]:
    result = []
    seen = set()
    for value in values:
        marker = json.dumps(json_safe(value), sort_keys=True)
        if marker not in seen:
            seen.add(marker)
            result.append(value)
    return result


def parse_json_object(content: str) -> Any:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content, flags=re.I)
        content = re.sub(r"\s*```$", "", content)
    return json.loads(content)


def json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return money_string(value)
    if isinstance(value, Path):
        return str(value)
    return value


def slugify(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", value).strip("_").lower()


def main(argv: Optional[list[str]] = None) -> int:
    try:
        load_dotenv()
        args = parse_args(argv)
        run(args)
        return 0
    except KeyboardInterrupt:
        print("Interrupted", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        if "--verbose" in (argv if argv is not None else sys.argv[1:]):
            raise
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
