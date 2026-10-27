"""Read a portfolio workbook and keep only dated bond positions."""

from __future__ import annotations

import calendar
import html
import math
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from typing import BinaryIO
from xml.etree import ElementTree as ET

import pandas as pd

BOND_TYPES = frozenset(
    {
        "Купонные облигации",
        "Корпоративные облигации",
        "Евроноты",
        "ГЦБ",
    }
)

REQUIRED_COLUMNS = (
    "Тикер",
    "Эмитент",
    "ISIN",
    "Тип ЦБ",
    "Валюта",
    "Количество",
    "Номинал",
    "Количество ЦБ по номиналу",
    "Текущая стоимость, в тенге",
    "Купонная ставка",
    "Дата ближайшей купонной выплаты",
    "Дата погашения",
)

STOP_LABELS = frozenset(
    {
        "Прочие обязательства",
        "Итого обязательства",
        "Итого обязателства",
        "Итого чистые активы",
        "Лимит на 1 эмитента",
    }
)

COLUMN_FIELDS = {
    "Тикер": "ticker",
    "Эмитент": "issuer",
    "ISIN": "isin",
    "Тип ЦБ": "security_type",
    "Валюта": "currency",
    "Количество": "quantity",
    "Номинал": "nominal",
    "Количество ЦБ по номиналу": "face_amount",
    "Текущая стоимость, в тенге": "value_kzt",
    "Купонная ставка": "coupon",
    "Дата ближайшей купонной выплаты": "next_coupon",
    "Дата погашения": "maturity",
}


@dataclass
class Portfolio:
    report_date: date | None
    fund_name: str | None
    fx_rates: dict[str, float]
    bonds: pd.DataFrame
    exceptions: pd.DataFrame


def add_months(day: date, months: int) -> date:
    """Add calendar months, clamping the day to the target month's length."""
    month_index = day.month - 1 + months
    year = day.year + month_index // 12
    month = month_index % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, min(day.day, last_day))


def filter_by_horizon(
    bonds: pd.DataFrame, report_date: date, months: int | None
) -> pd.DataFrame:
    """Keep issues that redeem on or after the report date and inside the horizon.

    ``months is None`` means every redemption on or after the report date.
    """
    if bonds.empty:
        return bonds.copy()
    maturity = pd.to_datetime(bonds["maturity"])
    report = pd.Timestamp(report_date)
    mask = maturity >= report
    if months is not None:
        mask = mask & (maturity <= pd.Timestamp(add_months(report_date, months)))
    window = bonds.loc[mask].copy()
    window["_maturity_sort"] = maturity.loc[mask].to_numpy()
    window = window.sort_values(["_maturity_sort", "issuer", "isin"]).drop(columns="_maturity_sort")
    return window.reset_index(drop=True)


def parse_portfolio(source: BinaryIO | str) -> Portfolio:
    """Parse the portfolio sheet from an ``.xlsx`` path or file object.

    Grouped positions are a green summary row plus hidden white lot rows.
    The summary already holds the position total, so only the first lot row
    is read (for the issuer, the type, and the clean ticker). The rest of
    the lots are skipped while the sheet is streamed.
    """
    rows = _load_rows(source)
    report_date, fund_name, lots, exceptions = _scan_rows(rows)
    bonds = _aggregate_lots(lots)
    rates = _fx_rates(rows, _header_row_index(rows))
    bonds = _to_kzt(bonds, rates)
    return Portfolio(
        report_date=report_date,
        fund_name=fund_name,
        fx_rates=rates,
        bonds=bonds,
        exceptions=exceptions,
    )


_SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PACKAGE_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CELL_RE = re.compile(r"<(?:[\w]+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:[\w]+:)?c>)", re.DOTALL)
_CELL_REF_RE = re.compile(r"""\br="([A-Z]+)\d+""")
_VALUE_RE = re.compile(r"<(?:[\w]+:)?v[^>]*>([^<]*)</(?:[\w]+:)?v>")
_INLINE_RE = re.compile(r"<(?:[\w]+:)?t[^>]*>([^<]*)</(?:[\w]+:)?t>")
_OUTLINE_RE = re.compile(r"""\boutlineLevel="(\d+)""")

_SUMMARY_FIELDS = (
    "nominal",
    "face_amount",
    "value_kzt",
    "currency",
    "coupon",
    "next_coupon",
    "maturity",
)


def _load_rows(source: BinaryIO | str) -> list[tuple]:
    if isinstance(source, str):
        archive = zipfile.ZipFile(source)
    else:
        data = source.read() if hasattr(source, "read") else source
        archive = zipfile.ZipFile(BytesIO(data))
    try:
        sheet_path = _portfolio_sheet_path(archive)
        strings = _shared_strings(archive)
        with archive.open(sheet_path) as sheet:
            return list(_iter_sheet_rows(sheet, strings))
    finally:
        archive.close()


def _portfolio_sheet_path(archive: zipfile.ZipFile) -> str:
    ns = {"m": _SPREADSHEET_NS, "r": _OFFICE_REL_NS, "p": _PACKAGE_NS}
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook.findall("m:sheets/m:sheet", ns)
    if not sheets:
        raise ValueError("В файле нет листов.")
    chosen = next(
        (sheet for sheet in sheets if _clean(sheet.get("name")) == "ПОРТФЕЛЬ"),
        sheets[0],
    )
    relation_id = chosen.get(f"{{{_OFFICE_REL_NS}}}id")
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels.findall("p:Relationship", ns):
        if rel.get("Id") == relation_id:
            target = rel.get("Target")
            break
    if not target:
        raise ValueError("Не найден лист портфеля.")
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    strings: list[str] = []
    tag = f"{{{_SPREADSHEET_NS}}}si"
    text_tag = f"{{{_SPREADSHEET_NS}}}t"
    for _event, elem in ET.iterparse(archive.open("xl/sharedStrings.xml"), events=("end",)):
        if elem.tag != tag:
            continue
        strings.append("".join(node.text or "" for node in elem.iter(text_tag)))
        elem.clear()
    return strings


def _iter_sheet_rows(sheet, strings: list[str]):
    """Read green totals, one matching white row, and any row that is a different position.

    Extra white lots of the same ISIN are not parsed. A white row with another
    ISIN is kept, because the file sometimes nests unrelated positions in the
    same outline group.
    """
    xml = sheet.read().decode("utf-8")
    isin_letters: str | None = None
    group_isin: str | None = None
    need_detail = False
    cursor = 0
    while True:
        start = xml.find("<row ", cursor)
        if start < 0:
            return
        tag_end = xml.find(">", start)
        if tag_end < 0:
            return
        attrs = xml[start + 4 : tag_end]
        if attrs.endswith("/"):
            cursor = tag_end + 1
            continue
        body_end = xml.find("</row>", tag_end)
        if body_end < 0:
            return
        level_match = _OUTLINE_RE.search(attrs)
        level = int(level_match.group(1)) if level_match else 0
        body = xml[tag_end + 1 : body_end]
        cursor = body_end + 6
        if level <= 1:
            row = _parse_row_body(body, strings)
            if isin_letters is None:
                header = _header_map(row)
                if header is not None:
                    isin_letters = _letters_from_index(header["ISIN"])
            if level == 1 and isin_letters:
                group_isin = _clean(_at(row, _column_index(isin_letters)))
                need_detail = group_isin is not None
            else:
                group_isin = None
                need_detail = False
            yield row
            continue
        same_group = (
            group_isin is not None
            and isin_letters is not None
            and _row_matches_isin(body, isin_letters, group_isin, strings)
        )
        if not same_group:
            yield _parse_row_body(body, strings)
        elif need_detail:
            need_detail = False
            yield _parse_row_body(body, strings)


def _parse_row_body(body: str, strings: list[str]) -> tuple:
    if not body or "<c" not in body:
        return ()
    cells: dict[int, object] = {}
    for match in _CELL_RE.finditer(body):
        attrs = match.group(1)
        ref = _CELL_REF_RE.search(attrs)
        if ref is None:
            continue
        value = _cell_value(attrs, match.group(2), strings)
        if value is None:
            continue
        cells[_column_index(ref.group(1))] = value
    if not cells:
        return ()
    row = [None] * (max(cells) + 1)
    for index, value in cells.items():
        row[index] = value
    return tuple(row)


def _cell_value(attrs: str, inner: str | None, strings: list[str]):
    if 't="s"' in attrs or "t='s'" in attrs:
        found = _VALUE_RE.search(inner or "")
        if found is None:
            return None
        index = int(found.group(1))
        return strings[index] if index < len(strings) else None
    if 't="inlineStr"' in attrs or "t='inlineStr'" in attrs:
        parts = _INLINE_RE.findall(inner or "")
        text = html.unescape("".join(parts))
        return text or None
    found = _VALUE_RE.search(inner or "")
    if found is None:
        return None
    text = found.group(1)
    try:
        if any(mark in text for mark in (".", "E", "e")):
            return float(text)
        return int(text)
    except ValueError:
        return text


def _row_matches_isin(body: str, letters: str, isin: str, strings: list[str]) -> bool:
    match = re.search(rf'<c r="{letters}\d+"([^>]*)(?:/>|>(.*?)</(?:[\w]+:)?c>)', body, re.DOTALL)
    if match is None:
        return False
    return _clean(_cell_value(match.group(1), match.group(2), strings)) == isin


def _letters_from_index(index: int) -> str:
    index += 1
    letters = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _column_index(letters: str) -> int:
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _scan_rows(rows: list[tuple]):
    preamble: list[tuple] = []
    columns: dict[str, int] | None = None
    lots: list[dict] = []
    exceptions: list[dict] = []
    blank_run = 0
    pending: dict | None = None
    for row in rows:
        if columns is None:
            columns = _header_map(row)
            if columns is None:
                preamble.append(row)
            continue
        if _is_stop_row(row):
            break
        raw = {field: _at(row, columns[label]) for label, field in COLUMN_FIELDS.items()}
        if _row_is_blank(raw):
            blank_run += 1
            if blank_run >= 25:
                break
            continue
        blank_run = 0
        if _is_group_summary(raw):
            pending = raw
            continue
        if pending is not None and _clean(raw["isin"]) == _clean(pending["isin"]):
            raw = _merge_summary(pending, raw)
            pending = None
        record = _bond_record(raw)
        if record is None:
            continue
        if record["maturity"] is None:
            exceptions.append(record)
        else:
            lots.append(record)
    if columns is None:
        raise ValueError("В файле нет таблицы ценных бумаг с ожидаемыми колонками.")
    report_date, fund_name = _report_date_from_preamble(preamble)
    return report_date, fund_name, pd.DataFrame(lots), _exceptions_frame(exceptions)


def _header_map(row: tuple) -> dict[str, int] | None:
    found: dict[str, int] = {}
    for index, value in enumerate(row):
        label = _clean(value)
        if label in REQUIRED_COLUMNS and label not in found:
            found[label] = index
    if all(name in found for name in REQUIRED_COLUMNS):
        return found
    return None


_CURRENCY_IN_TEXT = re.compile(r"[вВ]\s+([A-Za-z]{3})\s*$")


def _header_row_index(rows: list[tuple]) -> int:
    for index, row in enumerate(rows):
        if _header_map(row) is not None:
            return index
    return len(rows)


def _fx_rates(rows: list[tuple], header_index: int) -> dict[str, float]:
    """Rates from the cash block above the securities table. KZT per 1 unit of currency."""
    rates = {"KZT": 1.0}
    for row in rows[:header_index]:
        code = None
        code_index = None
        for index, value in enumerate(row):
            text = _clean(value)
            if not text:
                continue
            match = _CURRENCY_IN_TEXT.search(text)
            if match:
                code = match.group(1).upper()
                code_index = index
                break
        if code is None:
            continue
        rate = None
        for value in row[code_index + 1 :]:
            number = _parse_number(value)
            if number is not None and number > 0:
                rate = number
                break
        if rate is not None:
            rates.setdefault(code, rate)
    return rates


def _to_kzt(bonds: pd.DataFrame, rates: dict[str, float]) -> pd.DataFrame:
    if bonds.empty:
        bonds = bonds.copy()
        bonds["fx_rate"] = pd.Series(dtype=float)
        bonds["nominal_kzt"] = pd.Series(dtype=float)
        return bonds
    frame = bonds.copy()
    frame["fx_rate"] = frame["currency"].map(rates)
    frame["nominal_kzt"] = frame["nominal_amount"] * frame["fx_rate"]
    return frame


def _is_group_summary(raw: dict) -> bool:
    """Green row: position total, with the issuer and the type on the next white row."""
    if _clean(raw.get("issuer")) or not _clean(raw.get("isin")):
        return False
    return _parse_number(raw.get("face_amount")) is not None


def _merge_summary(summary: dict, detail: dict) -> dict:
    """Keep the green row's totals and the white row's issuer, type, and ticker."""
    merged = dict(detail)
    for key in _SUMMARY_FIELDS:
        value = summary.get(key)
        if value not in (None, ""):
            merged[key] = value
    merged["grouped"] = True
    return merged


def _report_date_from_preamble(preamble: list[tuple]) -> tuple[date | None, str | None]:
    for row in preamble:
        for index, value in enumerate(row):
            if _clean(value) != "за":
                continue
            report_date = None
            for candidate in row[index + 1 : index + 4]:
                report_date = _parse_date(candidate)
                if report_date is not None:
                    break
            fund_name = None
            for candidate in row[:index]:
                text = _clean(candidate)
                if text and _parse_date(candidate) is None:
                    fund_name = text
            return report_date, fund_name
    return None, None


def _bond_record(raw: dict) -> dict | None:
    security_type = _clean(raw["security_type"])
    issuer = _clean(raw["issuer"])
    isin = _clean(raw["isin"])
    if security_type not in BOND_TYPES or not issuer or not isin or _is_repo(issuer):
        return None
    return {
        "ticker": _clean(raw["ticker"]),
        "issuer": issuer,
        "isin": isin,
        "security_type": security_type,
        "currency": _clean(raw["currency"]),
        "quantity": _parse_number(raw["quantity"]),
        "nominal": _parse_number(raw["nominal"]),
        "face_amount": _parse_number(raw["face_amount"]),
        "value_kzt": _parse_number(raw["value_kzt"]),
        "coupon": _parse_number(raw["coupon"]),
        "next_coupon": _parse_date(raw["next_coupon"]),
        "maturity": _parse_date(raw["maturity"]),
        "nominal_unit": _parse_number(raw["nominal"]),
        "grouped": bool(raw.get("grouped")),
    }


def _at(row: tuple, index: int):
    if index >= len(row):
        return None
    return row[index]


def _is_stop_row(row: tuple) -> bool:
    for value in row[:6]:
        if _clean(value) in STOP_LABELS:
            return True
    return False


def _row_is_blank(raw: dict) -> bool:
    for value in raw.values():
        if _parse_date(value) is not None or _parse_number(value) is not None or _clean(value) is not None:
            return False
    return True


def _as_dates(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    return parsed.apply(lambda value: None if pd.isna(value) else value.date())


def _aggregate_lots(lots: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "ticker",
        "issuer",
        "isin",
        "security_type",
        "currency",
        "quantity",
        "nominal_amount",
        "value_kzt",
        "coupon",
        "next_coupon",
        "maturity",
    ]
    if lots.empty:
        return pd.DataFrame(columns=columns)

    frame = lots.copy()
    frame["nominal_amount"] = frame["face_amount"]
    aggregated = (
        frame.groupby("isin", as_index=False)
        .agg(
            ticker=("ticker", "first"),
            issuer=("issuer", "first"),
            security_type=("security_type", "first"),
            currency=("currency", "first"),
            quantity=("quantity", "sum"),
            nominal_amount=("nominal_amount", "sum"),
            value_kzt=("value_kzt", "sum"),
            coupon=("coupon", "first"),
            next_coupon=("next_coupon", "first"),
            maturity=("maturity", "first"),
            nominal_unit=("nominal_unit", "first"),
            grouped=("grouped", "any"),
        )
        .sort_values(["maturity", "issuer", "isin"])
        .reset_index(drop=True)
    )
    # The green row's «Количество» is not the sum of lots. Face divided by the
    # bond's own nominal is. Ungrouped rows already summed their quantities.
    grouped = aggregated["grouped"] & aggregated["nominal_unit"].notna() & (aggregated["nominal_unit"] != 0)
    aggregated.loc[grouped, "quantity"] = (
        aggregated.loc[grouped, "nominal_amount"] / aggregated.loc[grouped, "nominal_unit"]
    )
    aggregated["maturity"] = _as_dates(aggregated["maturity"])
    aggregated["next_coupon"] = _as_dates(aggregated["next_coupon"])
    return aggregated[columns]


def _exceptions_frame(records: list[dict]) -> pd.DataFrame:
    columns = [
        "ticker",
        "issuer",
        "isin",
        "security_type",
        "currency",
        "quantity",
        "value_kzt",
    ]
    if not records:
        return pd.DataFrame(columns=columns)
    frame = pd.DataFrame(records)
    return frame[columns].sort_values(["issuer", "isin"]).reset_index(drop=True)


def _is_repo(issuer: str) -> bool:
    return issuer.upper().startswith("REPO")


def _clean(value) -> str | None:
    if value is None or isinstance(value, (datetime, date)):
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, bool):
        return None
    text = str(value).replace("\xa0", " ").replace("\u202f", " ").strip()
    return text or None


def _parse_number(value) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = _clean(value)
    if text is None:
        return None
    text = text.replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _parse_date(value) -> date | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or not value.is_integer()):
            return None
        serial = int(value)
        if serial <= 0:
            return None
        try:
            return (datetime(1899, 12, 30) + pd.Timedelta(days=serial)).date()
        except (OverflowError, ValueError):
            return None
    text = _clean(value)
    if text is None:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None
