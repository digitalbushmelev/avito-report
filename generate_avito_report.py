# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import html
import json
import os
import posixpath
import re
import statistics
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable
from urllib import parse, request
from xml.etree import ElementTree as ET


def configure_ssl_cert_file() -> None:
    if os.environ.get("SSL_CERT_FILE"):
        return
    try:
        import certifi  # type: ignore
    except Exception:
        return
    os.environ["SSL_CERT_FILE"] = certifi.where()


configure_ssl_cert_file()


ROOT = Path(__file__).resolve().parent


def path_from_env(name: str, default: str) -> Path:
    raw = os.environ.get(name, "").strip()
    path = Path(raw or default)
    return path if path.is_absolute() else ROOT / path


OUTPUT_FILE = path_from_env("REPORT_OUTPUT_FILE", "avito_report.html")
INDEX_FILE = path_from_env("REPORT_INDEX_FILE", "index.html")
AVITO_DATA_DIR = path_from_env("REPORT_AVITO_DATA_DIR", "data/avito")
AVITO_CAMPAIGN_REGISTRY_FILE = path_from_env("REPORT_AVITO_CAMPAIGN_REGISTRY_FILE", "avito_campaigns.json")
REPORT_CAMPAIGN_NAME = os.environ.get("REPORT_CAMPAIGN_NAME", "Дружеский")
REPORT_CAMPAIGN_SLUG = "druzheskiy"
LEAD_FILES = (ROOT / "leads.csv", ROOT / "bitrix_leads.csv", ROOT / "data" / "leads.csv")
BITRIX_CONFIG_FILE = ROOT / "bitrix_config.json"
BITRIX_ENV_FILE = path_from_env("REPORT_BITRIX_ENV_FILE", "bitrix.env")
REPORT_BRAND_NAME = os.environ.get("REPORT_BRAND_NAME", "Авито Реклама")
REPORT_HEADING = os.environ.get("REPORT_HEADING", "Сводка Авито Реклама")
REPORT_EYEBROW = os.environ.get("REPORT_EYEBROW", "Внутренний отчет")
ALL_LEADS_KEY = "__all__"
BITRIX_AVITO_LEADS_CAMPAIGN = REPORT_CAMPAIGN_NAME
DEFAULT_BITRIX_UTM_SOURCE = "avito_media"
DEFAULT_BITRIX_UTM_MEDIUM = "banner"
DEFAULT_BITRIX_UTM_CAMPAIGNS = (
    "psk_druzheskiy-pv1750000",
    "psk_druzheskiy-7na7",
    "psk_druzheskiy-pv1500000",
)
DEFAULT_BITRIX_UTM_GROUPS = (
    {
        "source": "avito_media",
        "medium": "banner",
    },
    {
        "source": "avito_reklama",
        "medium": "cpc",
    },
)
DEFAULT_BITRIX_DEAL_CATEGORY_NAME = "Комфорт: прямые продажи"
DEFAULT_BITRIX_EXCLUDED_COMMENT_PHONES = ("71111111111",)
DEFAULT_BITRIX_EXCLUDED_STAGE_NAMES = ("Дубль. Создана новая сделка", "Наш сотрудник")
DEFAULT_BITRIX_UNTOUCHED_STAGE_NAMES = ("ЛИДГЕН", "Совершить первый контакт")

XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


def esc(value: object) -> str:
    return html.escape(text(value), quote=True)


def normalize_key(value: object) -> str:
    return text(value).casefold()


def normalize_utm_value(value: object) -> str:
    return re.sub(r"[-_]+", "_", normalize_key(value))


def normalize_phone(value: object) -> str:
    return re.sub(r"\D+", "", text(value))


def split_list(value: object) -> list[str]:
    return [item.strip() for item in text(value).split(",") if item.strip()]


def as_list(value: object) -> list[str]:
    if isinstance(value, (list, tuple, set)):
        return [text(item) for item in value if text(item)]
    return split_list(value)


def clean_utm_groups(groups: object) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for group in groups if isinstance(groups, (list, tuple)) else []:
        if not isinstance(group, dict):
            continue
        campaigns = (
            group.get("campaigns")
            or group.get("campaign_values")
            or group.get("utm_campaigns")
            or group.get("utm_campaign")
        )
        item = {
            "source": text(group.get("source") or group.get("utm_source") or group.get("utm_source_value")),
            "medium": text(group.get("medium") or group.get("utm_medium") or group.get("utm_medium_value")),
            "campaigns": as_list(campaigns),
        }
        if item["source"] or item["medium"] or item["campaigns"]:
            result.append(item)
    return result


def default_utm_groups() -> list[dict[str, object]]:
    return clean_utm_groups(DEFAULT_BITRIX_UTM_GROUPS)


def utm_groups_from_json(value: object) -> list[dict[str, object]]:
    raw = text(value)
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return []
    if isinstance(parsed, dict):
        parsed = parsed.get("groups", [])
    return clean_utm_groups(parsed)


def legacy_utm_groups(source: object, medium: object, campaigns: object) -> list[dict[str, object]]:
    return clean_utm_groups(
        [
            {
                "source": text(source) or DEFAULT_BITRIX_UTM_SOURCE,
                "medium": text(medium) or DEFAULT_BITRIX_UTM_MEDIUM,
                "campaigns": as_list(campaigns) or list(DEFAULT_BITRIX_UTM_CAMPAIGNS),
            }
        ]
    )


def to_float(value: object) -> float:
    raw = text(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not raw:
        return 0.0
    try:
        return float(raw)
    except ValueError:
        return 0.0


def to_int(value: object) -> int:
    return int(round(to_float(value)))


def number(value: object) -> str:
    return f"{to_float(value):,.0f}".replace(",", " ")


def money(value: object) -> str:
    return f"{to_float(value):,.0f}".replace(",", " ") + " ₽"


def decimal(value: object, digits: int = 1) -> str:
    return f"{to_float(value):.{digits}f}".replace(".", ",")


def pct(value: object, digits: int = 2) -> str:
    if value is None:
        return "нет данных"
    return decimal(value, digits) + "%"


def qualified_leads(total: dict[str, object]) -> float:
    return to_float(total.get("quality_leads"))


def appeal_count(total: dict[str, object]) -> float:
    return max(0.0, to_float(total.get("leads")) - qualified_leads(total))


def parse_date(value: object) -> date | None:
    raw = text(value)
    if not raw or raw.lower() == "итого":
        return None
    for candidate in (raw, raw[:10]):
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(candidate, fmt).date()
            except ValueError:
                pass
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    numeric = to_float(raw)
    if numeric:
        return date(1899, 12, 30) + timedelta(days=int(numeric))
    return None


def short_date(value: object) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d.%m") if parsed else text(value)


def display_date(value: object) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d.%m.%Y") if parsed else text(value)


def display_period(start: object, end: object) -> str:
    start_text = display_date(start)
    end_text = display_date(end)
    if start_text and end_text:
        return f"{start_text} — {end_text}"
    return start_text or end_text or "нет данных"


def col_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 1
    idx = 0
    for char in match.group(1):
        idx = idx * 26 + ord(char) - 64
    return idx


def relationship_target(target: str) -> str:
    normalized = target.lstrip("/")
    if normalized.startswith("xl/"):
        return normalized
    return posixpath.normpath("xl/" + normalized)


@dataclass
class Workbook:
    path: Path

    def sheets(self) -> dict[str, list[list[str]]]:
        with zipfile.ZipFile(self.path) as archive:
            shared = self._shared_strings(archive)
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

            result: dict[str, list[list[str]]] = {}
            for sheet in workbook.findall(f"{XLSX_NS}sheets/{XLSX_NS}sheet"):
                name = sheet.attrib["name"]
                sheet_path = relationship_target(rel_map[sheet.attrib[REL_ID]])
                root = ET.fromstring(archive.read(sheet_path))
                result[name] = self._rows(root, shared)
            return result

    @staticmethod
    def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return []
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        values = []
        for item in root.findall(f"{XLSX_NS}si"):
            values.append("".join(node.text or "" for node in item.findall(f".//{XLSX_NS}t")))
        return values

    @staticmethod
    def _rows(root: ET.Element, shared: list[str]) -> list[list[str]]:
        rows: list[list[str]] = []
        for row in root.findall(f"{XLSX_NS}sheetData/{XLSX_NS}row"):
            values: list[str] = []
            last_col = 0
            for cell in row.findall(f"{XLSX_NS}c"):
                current_col = col_index(cell.attrib.get("r", "A1"))
                while last_col + 1 < current_col:
                    values.append("")
                    last_col += 1

                value_type = cell.attrib.get("t")
                raw_value = cell.find(f"{XLSX_NS}v")
                if value_type == "s" and raw_value is not None:
                    value = shared[int(raw_value.text or "0")]
                elif value_type == "inlineStr":
                    value = "".join(node.text or "" for node in cell.findall(f".//{XLSX_NS}t"))
                elif raw_value is not None:
                    value = raw_value.text or ""
                else:
                    value = ""
                values.append(value)
                last_col = current_col
            rows.append(values)
        return rows


@dataclass
class Export:
    kind: str
    path: Path
    info: dict[str, str]
    sheets: dict[str, list[list[str]]]
    display_name: str


@dataclass
class LeadData:
    enabled: bool
    source: str
    status: str
    by_campaign: dict[str, dict[str, int]]
    quality_by_campaign: dict[str, dict[str, dict[str, int]]] = field(default_factory=dict)
    stage_summary: dict[str, dict[str, object]] = field(default_factory=dict)
    by_utm_campaign: dict[str, dict[str, int]] = field(default_factory=dict)


def read_info(rows: list[list[str]]) -> dict[str, str]:
    info: dict[str, str] = {}
    for row in rows:
        if not row:
            continue
        label = text(row[0])
        if label.startswith("Кампания"):
            info["campaign_id"] = text(row[1]) if len(row) > 1 else ""
            info["campaign_name"] = text(row[2]) if len(row) > 2 else ""
        elif label == "Тип кампании":
            info["campaign_type"] = text(row[1]) if len(row) > 1 else ""
        elif label == "Модель оплаты":
            info["payment_model"] = text(row[1]) if len(row) > 1 else ""
        elif label.startswith("Рекламодатель"):
            info["advertiser_inn"] = text(row[1]) if len(row) > 1 else ""
            info["advertiser_name"] = text(row[2]) if len(row) > 2 else ""
        elif label.startswith("Период отчета"):
            info["period_start"] = text(row[1]) if len(row) > 1 else ""
            info["period_end"] = text(row[2]) if len(row) > 2 else ""
        elif label.startswith("Дата выгрузки"):
            info["export_date"] = text(row[1]) if len(row) > 1 else ""
    return info


def classify_workbook(path: Path, sheets: dict[str, list[list[str]]]) -> str:
    name = path.name.casefold()
    if "демография" in name:
        return "demography"
    if "география" in name:
        return "geography"
    sheet_names = " ".join(sheets).casefold()
    if " пол" in sheet_names or " возраст" in sheet_names or " доход" in sheet_names:
        return "demography"
    if " гео" in sheet_names:
        return "geography"
    return "performance"


def campaign_folder_name(path: Path) -> str:
    parent = path.parent
    return parent.name if parent != ROOT else ""


def discover_exports() -> dict[str, dict[str, Export]]:
    grouped: dict[str, dict[str, Export]] = {}
    for path in ROOT.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        sheets = Workbook(path).sheets()
        info = read_info(sheets.get("Инфо", []))
        display_name = campaign_folder_name(path)
        if display_name:
            info["display_name"] = display_name
        campaign_id = info.get("campaign_id") or campaign_id_from_filename(path.name)
        if not campaign_id:
            continue
        kind = classify_workbook(path, sheets)
        export = Export(kind=kind, path=path, info=info, sheets=sheets, display_name=display_name)
        current = grouped.setdefault(campaign_id, {}).get(kind)
        if current is None or path.stat().st_mtime > current.path.stat().st_mtime:
            grouped[campaign_id][kind] = export
    return grouped


def campaign_id_from_filename(name: str) -> str:
    match = re.search(r"\((\d{5,})\)", name)
    return match.group(1) if match else ""


def row_value(row: list[str], index: dict[str, int], name: str) -> str:
    pos = index.get(name)
    if pos is None or pos >= len(row):
        return ""
    return row[pos]


def empty_total() -> dict[str, float]:
    return {
        "impressions": 0.0,
        "clicks": 0.0,
        "spend": 0.0,
        "bonus_spend": 0.0,
        "ctr": 0.0,
        "cpc": 0.0,
        "cpm": 0.0,
    }


def metric_record(row: list[str], index: dict[str, int]) -> dict[str, float]:
    impressions = to_int(row_value(row, index, "Показы"))
    clicks = to_int(row_value(row, index, "Клики"))
    spend = to_float(row_value(row, index, "Расход, Руб."))
    bonus = to_float(row_value(row, index, "Расход бонусами, Руб."))
    return complete_rates(
        {
            "impressions": float(impressions),
            "clicks": float(clicks),
            "spend": spend,
            "bonus_spend": bonus,
        }
    )


def complete_rates(total: dict[str, float]) -> dict[str, float]:
    impressions = total.get("impressions", 0.0)
    clicks = total.get("clicks", 0.0)
    spend = total.get("spend", 0.0)
    total["ctr"] = clicks / impressions * 100 if impressions else 0.0
    total["cpc"] = spend / clicks if clicks else 0.0
    total["cpm"] = spend / impressions * 1000 if impressions else 0.0
    return total


def parse_metric_sheet(rows: list[list[str]]) -> tuple[list[dict[str, object]], dict[str, float]]:
    if not rows:
        return [], empty_total()
    headers = [text(value) for value in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    daily: list[dict[str, object]] = []
    explicit_total: dict[str, float] | None = None

    for row in rows[1:]:
        if not row:
            continue
        first = text(row[0])
        if not first:
            continue
        metrics = metric_record(row, index)
        if first.casefold() == "итого":
            explicit_total = metrics
            continue
        day = parse_date(first)
        if not day:
            continue
        daily.append({"date": day.isoformat(), **metrics})

    total = explicit_total if explicit_total is not None else sum_metric_rows(daily)
    return daily, complete_rates(total)


def sum_metric_rows(rows: Iterable[dict[str, object]]) -> dict[str, float]:
    total = empty_total()
    for row in rows:
        total["impressions"] += to_float(row.get("impressions"))
        total["clicks"] += to_float(row.get("clicks"))
        total["spend"] += to_float(row.get("spend"))
        total["bonus_spend"] += to_float(row.get("bonus_spend"))
    return complete_rates(total)


def parse_entity(sheet_name: str, prefix: str) -> tuple[int, str]:
    match = re.search(rf"{prefix}\s+(\d+)\((\d+)\)", sheet_name)
    if not match:
        return 0, ""
    return int(match.group(1)), match.group(2)


def summarize_entity(sheet_name: str, prefix: str, rows: list[list[str]]) -> dict[str, object]:
    daily, total = parse_metric_sheet(rows)
    index, entity_id = parse_entity(sheet_name, prefix)
    return {
        "sheet": sheet_name,
        "index": index,
        "id": entity_id,
        "active_days": len(active_rows(daily)),
        "daily": daily,
        **total,
    }


def parse_distribution(rows: list[list[str]], label_name: str = "") -> list[dict[str, object]]:
    if not rows:
        return []
    headers = [text(value) for value in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    label_col = index.get(label_name) if label_name else 0
    clicks_col = index.get("Клики", len(headers) - 1)
    result = []
    for row in rows[1:]:
        if not row:
            continue
        label = text(row[label_col]) if label_col is not None and label_col < len(row) else ""
        if not label:
            continue
        clicks = to_int(row[clicks_col]) if clicks_col < len(row) else 0
        result.append({"label": label, "clicks": clicks})
    return result


def parse_age(rows: list[list[str]]) -> dict[str, object]:
    if not rows:
        return {"rows": [], "totals": [], "by_gender": {}}
    headers = [text(value) for value in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    by_age: dict[str, int] = {}
    by_gender: dict[str, dict[str, int]] = {}
    raw_rows = []
    for row in rows[1:]:
        age = text(row[index.get("Возраст", 0)]) if index.get("Возраст", 0) < len(row) else ""
        gender = text(row[index.get("Пол", 1)]).strip() if index.get("Пол", 1) < len(row) else ""
        clicks = to_int(row[index.get("Клики", 2)]) if index.get("Клики", 2) < len(row) else 0
        if not age:
            continue
        raw_rows.append({"age": age, "gender": gender, "clicks": clicks})
        by_age[age] = by_age.get(age, 0) + clicks
        by_gender.setdefault(age, {})[gender] = by_gender.setdefault(age, {}).get(gender, 0) + clicks

    order = ["<17", "18-24", "25-34", "35-44", "45-54", "55-64", ">65", "Не определено"]
    totals = [{"label": age, "clicks": by_age.get(age, 0)} for age in order if age in by_age]
    return {"rows": raw_rows, "totals": totals, "by_gender": by_gender}


def add_shares(items: list[dict[str, object]], total: float | None = None) -> list[dict[str, object]]:
    base = total if total is not None else sum(to_float(item.get("clicks")) for item in items)
    enriched = []
    for item in items:
        clicks = to_float(item.get("clicks"))
        enriched.append({**item, "share": clicks / base * 100 if base else 0.0})
    return enriched


def active_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    return [
        row
        for row in rows
        if to_float(row.get("impressions")) > 0 or to_float(row.get("clicks")) > 0 or to_float(row.get("spend")) > 0
    ]


def longest_zero_gap(rows: list[dict[str, object]]) -> int:
    best = 0
    current = 0
    for row in rows:
        if row in active_rows([row]):
            current = 0
        else:
            current += 1
            best = max(best, current)
    return best


def in_date_range(day: date | None, date_from: date | None, date_to: date | None) -> bool:
    if day and date_from and day < date_from:
        return False
    if day and date_to and day > date_to:
        return False
    return True


def add_leads(target: dict[str, dict[str, int]], campaign_key: str, date_key: str, count: int) -> None:
    if not campaign_key:
        return
    target.setdefault(campaign_key, {})[date_key] = target.setdefault(campaign_key, {}).get(date_key, 0) + count


def add_lead_quality(
    target: dict[str, dict[str, dict[str, int]]],
    campaign_key: str,
    date_key: str,
    quality_key: str,
    count: int,
) -> None:
    if not campaign_key:
        return
    bucket = target.setdefault(campaign_key, {}).setdefault(date_key, {"quality": 0, "bad": 0})
    bucket[quality_key] = bucket.get(quality_key, 0) + count


def merge_lead_quality(
    source: dict[str, dict[str, dict[str, int]]],
    campaign_keys: Iterable[str],
    fallback_all: bool = False,
) -> dict[str, dict[str, int]]:
    result: dict[str, dict[str, int]] = {}
    keys = [key for key in campaign_keys if key]
    for key in keys:
        for day, quality in source.get(key, {}).items():
            bucket = result.setdefault(day, {"quality": 0, "bad": 0})
            bucket["quality"] += to_int(quality.get("quality"))
            bucket["bad"] += to_int(quality.get("bad"))
    if not result and fallback_all:
        for day, quality in source.get(ALL_LEADS_KEY, {}).items():
            bucket = result.setdefault(day, {"quality": 0, "bad": 0})
            bucket["quality"] += to_int(quality.get("quality"))
            bucket["bad"] += to_int(quality.get("bad"))
    return result


def add_stage_summary(
    target: dict[str, dict[str, object]],
    stage_id: str,
    stage_name: str,
    quality_key: str,
    date_key: str,
    count: int,
) -> None:
    key = stage_id or stage_name or "unknown"
    bucket = target.setdefault(
        key,
        {
            "id": stage_id,
            "name": stage_name or stage_id or "Без стадии",
            "quality": quality_key,
            "count": 0,
        },
    )
    bucket["count"] = to_int(bucket.get("count")) + count
    by_day = bucket.setdefault("by_day", {})
    if isinstance(by_day, dict):
        by_day[date_key] = to_int(by_day.get(date_key)) + count


def read_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if path.exists():
        for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            value = value.strip().strip('"').strip("'")
            values[key.strip()] = value
    for key, value in os.environ.items():
        if key.startswith(("BITRIX_", "REPORT_")):
            values[key] = value
    return values


def encode_bitrix_params(params: object, prefix: str = "") -> list[tuple[str, str | int]]:
    if isinstance(params, dict):
        result: list[tuple[str, str | int]] = []
        for key, value in params.items():
            nested_prefix = f"{prefix}[{key}]" if prefix else text(key)
            result.extend(encode_bitrix_params(value, nested_prefix))
        return result
    if isinstance(params, (list, tuple, set)):
        result = []
        for value in params:
            result.extend(encode_bitrix_params(value, prefix + "[]"))
        return result
    return [(prefix, params if isinstance(params, int) else text(params))]


def bitrix_request(webhook_url: str, method: str, params: dict[str, object] | None = None) -> tuple[dict[str, object], str | None]:
    endpoint = webhook_url.rstrip("/") + "/" + method + ".json"
    payload = parse.urlencode(encode_bitrix_params(params or {})).encode("utf-8")
    req = request.Request(endpoint, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        return {}, str(exc)

    if data.get("error"):
        return {}, text(data.get("error_description") or data.get("error"))
    return data, None


def bitrix_list(webhook_url: str, method: str, select_fields: list[str], filters: dict[str, str]) -> tuple[list[dict[str, object]], str | None]:
    rows: list[dict[str, object]] = []
    start: int | str = 0

    while True:
        params: dict[str, object] = {
            "start": start,
            "select": list(dict.fromkeys(field for field in select_fields if field)),
            "filter": {key: value for key, value in filters.items() if value != ""},
            "order": {"ID": "ASC"},
        }

        data, error = bitrix_request(webhook_url, method, params)
        if error:
            return [], error

        result = data.get("result", [])
        page_rows = result.get("items", []) if isinstance(result, dict) else result
        rows.extend(page_rows if isinstance(page_rows, list) else [])

        next_start = data.get("next")
        if next_start is None:
            break
        start = next_start

    return rows, None


def extract_bitrix_items(result: object, *keys: str) -> list[dict[str, object]]:
    if isinstance(result, list):
        return [item for item in result if isinstance(item, dict)]
    if not isinstance(result, dict):
        return []
    for key in keys:
        items = result.get(key)
        if isinstance(items, list):
            return [item for item in items if isinstance(item, dict)]
    return []


def bitrix_item_value(item: dict[str, object], *keys: str) -> object:
    for key in keys:
        if key in item and item[key] is not None:
            return item[key]
    return ""


def resolve_bitrix_deal_category_id(webhook_url: str, category_name: str) -> tuple[str, str | None]:
    target = normalize_key(category_name)
    if not target:
        return "", None

    methods = [
        ("crm.category.list", {"entityTypeId": 2}, ("categories", "items")),
        ("crm.dealcategory.list", {}, ("categories", "items")),
    ]
    last_error = ""
    for method, params, item_keys in methods:
        data, error = bitrix_request(webhook_url, method, params)
        if error:
            last_error = error
            continue
        result = data.get("result", [])
        for item in extract_bitrix_items(result, *item_keys):
            name = text(bitrix_item_value(item, "name", "NAME"))
            if normalize_key(name) == target:
                return text(bitrix_item_value(item, "id", "ID")), None

    if last_error:
        return "", last_error
    return "", f"воронка не найдена: {category_name}"


def resolve_bitrix_deal_stage_ids(
    webhook_url: str,
    category_id: str,
    stage_names: Iterable[str],
) -> tuple[set[str], list[str], str | None]:
    targets = {normalize_key(name): text(name) for name in stage_names if text(name)}
    if not targets:
        return set(), [], None

    methods = [
        ("crm.status.list", {"filter": {"ENTITY_ID": f"DEAL_STAGE_{category_id}"}}, ("items", "stages", "statuses")),
        ("crm.dealcategory.stage.list", {"id": category_id}, ("items", "stages", "statuses")),
    ]
    last_error = ""
    matched: set[str] = set()
    matched_names: set[str] = set()
    for method, params, item_keys in methods:
        data, error = bitrix_request(webhook_url, method, params)
        if error:
            last_error = error
            continue
        result = data.get("result", [])
        for item in extract_bitrix_items(result, *item_keys):
            stage_id = text(bitrix_item_value(item, "STATUS_ID", "statusId", "ID", "id"))
            stage_name = text(bitrix_item_value(item, "NAME", "name"))
            normalized_name = normalize_key(stage_name)
            if normalized_name in targets and stage_id:
                matched.add(stage_id)
                matched_names.add(normalized_name)
        if len(matched_names) == len(targets):
            break

    missing = [original for key, original in targets.items() if key not in matched_names]
    if last_error and not matched:
        return set(), list(targets.values()), last_error
    return matched, missing, None


def bitrix_deal_stage_map(webhook_url: str, category_id: str) -> tuple[dict[str, dict[str, str]], str | None]:
    methods = [
        ("crm.status.list", {"filter": {"ENTITY_ID": f"DEAL_STAGE_{category_id}"}}, ("items", "stages", "statuses")),
        ("crm.dealcategory.stage.list", {"id": category_id}, ("items", "stages", "statuses")),
    ]
    last_error = ""
    for method, params, item_keys in methods:
        data, error = bitrix_request(webhook_url, method, params)
        if error:
            last_error = error
            continue
        result = data.get("result", [])
        stages: dict[str, dict[str, str]] = {}
        for item in extract_bitrix_items(result, *item_keys):
            stage_id = text(bitrix_item_value(item, "STATUS_ID", "statusId", "ID", "id"))
            if not stage_id:
                continue
            stages[stage_id] = {
                "name": text(bitrix_item_value(item, "NAME", "name")),
                "semantics": text(bitrix_item_value(item, "SEMANTICS", "semantics")).upper(),
            }
        if stages:
            return stages, None
    return {}, last_error or None


def add_campaign_date_rule(target: list[dict[str, object]], campaign_key: object, rule: object) -> None:
    if not isinstance(rule, dict):
        return
    key = normalize_key(rule.get("campaign_id") or rule.get("id") or rule.get("campaign_name") or rule.get("name") or campaign_key)
    if not key:
        return
    date_from = parse_date(rule.get("date_from") or rule.get("from") or rule.get("start"))
    date_to = parse_date(rule.get("date_to") or rule.get("to") or rule.get("end"))
    if not date_from and not date_to:
        return
    target.append({"campaign_key": key, "date_from": date_from, "date_to": date_to})


def campaign_date_rules_from_json(value: object) -> list[dict[str, object]]:
    raw = text(value)
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return []

    rules: list[dict[str, object]] = []
    if isinstance(parsed, dict):
        for campaign_key, rule in parsed.items():
            add_campaign_date_rule(rules, campaign_key, rule)
    else:
        for item in parsed if isinstance(parsed, list) else []:
            add_campaign_date_rule(rules, "", item)
    return rules


def campaign_date_rules_from_config(value: object) -> list[dict[str, object]]:
    rules: list[dict[str, object]] = []
    if isinstance(value, dict):
        for campaign_key, rule in value.items():
            add_campaign_date_rule(rules, campaign_key, rule)
    else:
        for item in value if isinstance(value, list) else []:
            add_campaign_date_rule(rules, "", item)
    return rules


def campaign_date_rules_from_registry() -> list[dict[str, object]]:
    if not AVITO_CAMPAIGN_REGISTRY_FILE.exists():
        return []
    try:
        parsed = json.loads(AVITO_CAMPAIGN_REGISTRY_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    items = parsed.get("campaigns") if isinstance(parsed, dict) else parsed
    rules: list[dict[str, object]] = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        rule = {
            "campaign_id": item.get("id"),
            "campaign_name": item.get("name"),
            "date_from": item.get("lead_date_from"),
            "date_to": item.get("lead_date_to"),
        }
        add_campaign_date_rule(rules, "", rule)
    return rules


def campaign_registry_meta() -> dict[str, dict[str, object]]:
    if not AVITO_CAMPAIGN_REGISTRY_FILE.exists():
        return {}
    try:
        parsed = json.loads(AVITO_CAMPAIGN_REGISTRY_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    items = parsed.get("campaigns") if isinstance(parsed, dict) else parsed
    result: dict[str, dict[str, object]] = {}
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        campaign_id = text(item.get("id"))
        if not campaign_id:
            continue
        result[campaign_id] = {
            "id": campaign_id,
            "name": text(item.get("name") or campaign_id),
            "lead_keys": as_list(item.get("lead_keys")),
            "lead_date_from": text(item.get("lead_date_from") or item.get("date_from")),
            "lead_date_to": text(item.get("lead_date_to") or item.get("date_to")),
        }
    return result


def campaign_key_for_date(rules: list[dict[str, object]], day: date | None) -> str:
    if not day:
        return ""
    for rule in rules:
        date_from = rule.get("date_from")
        date_to = rule.get("date_to")
        if isinstance(date_from, date) and day < date_from:
            continue
        if isinstance(date_to, date) and day > date_to:
            continue
        return text(rule.get("campaign_key"))
    return ""


def bitrix_settings_from_env() -> dict[str, object] | None:
    env = read_env_file(BITRIX_ENV_FILE)
    if not env:
        return None

    entity_type = text(env.get("BITRIX_ENTITY_TYPE") or "deal").lower()
    if entity_type == "lead":
        entity_type = "deal"
    method = text(env.get("BITRIX_METHOD") or f"crm.{entity_type}.list")
    campaigns = text(env.get("BITRIX_UTM_CAMPAIGNS"))
    utm_groups = utm_groups_from_json(env.get("BITRIX_UTM_GROUPS_JSON"))
    if not utm_groups:
        has_legacy_utm = any(
            text(env.get(key))
            for key in ("BITRIX_UTM_SOURCE_VALUE", "BITRIX_UTM_MEDIUM_VALUE", "BITRIX_UTM_CAMPAIGNS")
        )
        utm_groups = (
            legacy_utm_groups(
                env.get("BITRIX_UTM_SOURCE_VALUE"),
                env.get("BITRIX_UTM_MEDIUM_VALUE"),
                campaigns,
            )
            if has_legacy_utm
            else default_utm_groups()
        )
    excluded_phones = split_list(env.get("BITRIX_EXCLUDED_COMMENT_PHONES") or env.get("BITRIX_EXCLUDED_PHONE"))
    excluded_stage_names = split_list(env.get("BITRIX_EXCLUDED_STAGE_NAMES"))
    untouched_stage_names = split_list(env.get("BITRIX_UNTOUCHED_STAGE_NAMES"))
    campaign_date_rules = (
        campaign_date_rules_from_json(env.get("BITRIX_CAMPAIGN_DATE_MAP_JSON") or env.get("REPORT_CAMPAIGN_DATE_MAP_JSON"))
        or campaign_date_rules_from_registry()
    )
    return {
        "source": "bitrix.env",
        "webhook_url": text(env.get("BITRIX_WEBHOOK_URL")),
        "method": method,
        "date_field": text(env.get("BITRIX_DATE_FIELD") or "DATE_CREATE"),
        "utm_source_field": text(env.get("BITRIX_UTM_SOURCE_FIELD") or "UTM_SOURCE"),
        "utm_medium_field": text(env.get("BITRIX_UTM_MEDIUM_FIELD") or "UTM_MEDIUM"),
        "campaign_field": text(env.get("BITRIX_UTM_CAMPAIGN_FIELD") or "UTM_CAMPAIGN"),
        "utm_groups": utm_groups,
        "deal_category_id": text(env.get("BITRIX_DEAL_CATEGORY_ID") or env.get("BITRIX_CATEGORY_ID") or ""),
        "deal_category_name": text(
            env.get("BITRIX_DEAL_CATEGORY_NAME")
            or env.get("BITRIX_CATEGORY_NAME")
            or DEFAULT_BITRIX_DEAL_CATEGORY_NAME
        ),
        "stage_field": text(env.get("BITRIX_STAGE_FIELD") or "STAGE_ID"),
        "comment_field": text(env.get("BITRIX_COMMENT_FIELD") or "COMMENTS"),
        "excluded_stage_ids": split_list(env.get("BITRIX_EXCLUDED_STAGE_IDS")),
        "excluded_stage_names": excluded_stage_names or list(DEFAULT_BITRIX_EXCLUDED_STAGE_NAMES),
        "untouched_stage_names": untouched_stage_names or list(DEFAULT_BITRIX_UNTOUCHED_STAGE_NAMES),
        "excluded_comment_phones": excluded_phones or list(DEFAULT_BITRIX_EXCLUDED_COMMENT_PHONES),
        "lead_count_field": text(env.get("BITRIX_LEAD_COUNT_FIELD") or ""),
        "campaign_date_rules": campaign_date_rules,
    }


def bitrix_settings_from_config() -> dict[str, object] | None:
    if not BITRIX_CONFIG_FILE.exists():
        return None
    config = json.loads(BITRIX_CONFIG_FILE.read_text(encoding="utf-8"))
    if not config.get("enabled"):
        return None
    campaign_date_rules = campaign_date_rules_from_config(
        config.get("campaign_date_map") or config.get("campaign_date_rules")
    ) or campaign_date_rules_from_registry()
    return {
        "source": "bitrix_config.json",
        "webhook_url": text(config.get("webhook_url")),
        "method": text(config.get("method", "crm.deal.list")),
        "date_field": text(config.get("date_field", "DATE_CREATE")),
        "utm_source_field": text(config.get("utm_source_field", "UTM_SOURCE")),
        "utm_medium_field": text(config.get("utm_medium_field", "UTM_MEDIUM")),
        "campaign_field": text(config.get("campaign_field") or config.get("utm_campaign_field") or "UTM_CAMPAIGN"),
        "utm_groups": clean_utm_groups(config.get("utm_groups")) or legacy_utm_groups(
            config.get("utm_source_value"),
            config.get("utm_medium_value"),
            config.get("campaign_values") or list(DEFAULT_BITRIX_UTM_CAMPAIGNS),
        ),
        "deal_category_id": text(config.get("deal_category_id") or config.get("category_id") or ""),
        "deal_category_name": text(
            config.get("deal_category_name")
            or config.get("category_name")
            or DEFAULT_BITRIX_DEAL_CATEGORY_NAME
        ),
        "stage_field": text(config.get("stage_field") or "STAGE_ID"),
        "comment_field": text(config.get("comment_field") or "COMMENTS"),
        "excluded_stage_ids": config.get("excluded_stage_ids") or [],
        "excluded_stage_names": config.get("excluded_stage_names") or list(DEFAULT_BITRIX_EXCLUDED_STAGE_NAMES),
        "untouched_stage_names": config.get("untouched_stage_names") or list(DEFAULT_BITRIX_UNTOUCHED_STAGE_NAMES),
        "excluded_comment_phones": config.get("excluded_comment_phones") or list(DEFAULT_BITRIX_EXCLUDED_COMMENT_PHONES),
        "lead_count_field": text(config.get("lead_count_field")),
        "campaign_date_rules": campaign_date_rules,
    }


def load_leads_from_csv(date_from: date | None = None, date_to: date | None = None) -> LeadData | None:
    file_path = next((path for path in LEAD_FILES if path.exists()), None)
    if not file_path:
        return None

    by_campaign: dict[str, dict[str, int]] = {}
    quality_by_campaign: dict[str, dict[str, dict[str, int]]] = {}
    with file_path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames:
            return LeadData(False, file_path.name, "CSV с лидами пустой", {})
        field_map = {normalize_key(name): name for name in reader.fieldnames}

        def field(*names: str) -> str | None:
            for name in names:
                if normalize_key(name) in field_map:
                    return field_map[normalize_key(name)]
            return None

        campaign_field = field("campaign_id", "campaign", "campaign_name", "utm_campaign", "id кампании", "кампания")
        date_field = field("date", "created_at", "created_date", "дата", "дата создания")
        leads_field = field("leads", "lead_count", "лиды", "количество лидов")
        if not campaign_field:
            return LeadData(False, file_path.name, "В CSV нет campaign_id/campaign_name", {})

        for row in reader:
            campaign_key = normalize_key(row.get(campaign_field))
            if not campaign_key:
                continue
            day = parse_date(row.get(date_field)) if date_field else None
            if not in_date_range(day, date_from, date_to):
                continue
            date_key = day.isoformat() if day else "unknown"
            lead_count = to_int(row.get(leads_field)) if leads_field else 1
            add_leads(by_campaign, campaign_key, date_key, lead_count)
            add_leads(by_campaign, ALL_LEADS_KEY, date_key, lead_count)
            add_lead_quality(quality_by_campaign, campaign_key, date_key, "quality", lead_count)
            add_lead_quality(quality_by_campaign, ALL_LEADS_KEY, date_key, "quality", lead_count)

    return LeadData(True, file_path.name, f"Лиды загружены из {file_path.name}", by_campaign, quality_by_campaign)


def load_leads_from_bitrix(date_from: date | None = None, date_to: date | None = None) -> LeadData | None:
    settings = bitrix_settings_from_env() or bitrix_settings_from_config()
    if not settings:
        return None

    webhook_url = text(settings.get("webhook_url"))
    method = text(settings.get("method"))
    date_field = text(settings.get("date_field"))
    utm_source_field = text(settings.get("utm_source_field"))
    utm_medium_field = text(settings.get("utm_medium_field"))
    campaign_field = text(settings.get("campaign_field"))
    utm_groups = clean_utm_groups(settings.get("utm_groups")) or default_utm_groups()
    deal_category_id = text(settings.get("deal_category_id"))
    deal_category_name = text(settings.get("deal_category_name"))
    stage_field = text(settings.get("stage_field") or "STAGE_ID")
    comment_field = text(settings.get("comment_field") or "COMMENTS")
    excluded_stage_ids = {text(value) for value in as_list(settings.get("excluded_stage_ids", [])) if text(value)}
    excluded_stage_names = as_list(settings.get("excluded_stage_names", []))
    untouched_stage_names = {normalize_key(value) for value in as_list(settings.get("untouched_stage_names", [])) if text(value)}
    excluded_comment_phones = {
        normalize_phone(value)
        for value in as_list(settings.get("excluded_comment_phones", []))
        if normalize_phone(value)
    }
    lead_count_field = text(settings.get("lead_count_field"))
    campaign_date_rules = [
        rule for rule in settings.get("campaign_date_rules", []) if isinstance(rule, dict)
    ]
    source = text(settings.get("source") or "Bitrix")

    missing = [
        name
        for name, value in (
            ("BITRIX_WEBHOOK_URL", webhook_url),
            ("BITRIX_DATE_FIELD", date_field),
            ("BITRIX_UTM_SOURCE_FIELD", utm_source_field),
            ("BITRIX_UTM_MEDIUM_FIELD", utm_medium_field),
        )
        if not value
    ]
    if missing:
        return LeadData(False, source, "В Bitrix-настройках не заполнено: " + ", ".join(missing), {})

    if deal_category_name and not deal_category_id:
        deal_category_id, category_error = resolve_bitrix_deal_category_id(webhook_url, deal_category_name)
        if category_error:
            return LeadData(False, "Bitrix", f"Bitrix: {category_error}", {})
    if not deal_category_id:
        return LeadData(False, "Bitrix", "В Bitrix-настройках не заполнено: BITRIX_DEAL_CATEGORY_ID", {})
    if not utm_groups:
        return LeadData(False, "Bitrix", "В Bitrix-настройках не заполнены UTM-группы", {})

    missing_stage_names: list[str] = []
    if excluded_stage_names:
        resolved_stage_ids, missing_stage_names, stage_error = resolve_bitrix_deal_stage_ids(
            webhook_url,
            deal_category_id,
            excluded_stage_names,
        )
        if stage_error:
            return LeadData(False, "Bitrix", f"Bitrix не отдал стадии воронки: {stage_error}", {})
        excluded_stage_ids.update(resolved_stage_ids)

    stage_map, stage_map_error = bitrix_deal_stage_map(webhook_url, deal_category_id)

    select_fields = [
        "ID",
        "CATEGORY_ID",
        stage_field,
        comment_field,
        date_field,
        utm_source_field,
        utm_medium_field,
        campaign_field,
        lead_count_field,
    ]

    by_campaign: dict[str, dict[str, int]] = {}
    by_utm_campaign: dict[str, dict[str, int]] = {}
    quality_by_campaign: dict[str, dict[str, dict[str, int]]] = {}
    stage_summary: dict[str, dict[str, object]] = {}
    matched_count = 0
    excluded_count = 0
    excluded_stage_only = 0
    excluded_phone_only = 0
    excluded_stage_and_phone = 0
    seen_deals: set[str] = set()
    utm_campaign_count = 0
    unknown_utm_campaign_count = 0
    date_mapped_count = 0
    for group in utm_groups:
        filters = {"=CATEGORY_ID": deal_category_id}
        if text(group.get("source")):
            filters[f"={utm_source_field}"] = text(group.get("source"))
        if utm_medium_field and text(group.get("medium")):
            filters[f"={utm_medium_field}"] = text(group.get("medium"))
        if date_from:
            filters[f">={date_field}"] = date_from.isoformat()
        if date_to:
            filters[f"<={date_field}"] = date_to.isoformat() + " 23:59:59"

        rows, error = bitrix_list(webhook_url, method, select_fields, filters)
        if error:
            return LeadData(False, "Bitrix", f"Bitrix не ответил: {error}", {})

        for row in rows:
            utm_campaign_key = normalize_utm_value(row.get(campaign_field))
            deal_id = text(row.get("ID")) or f"{utm_campaign_key}:{text(row.get(date_field))}:{text(row.get(comment_field))}"
            if deal_id in seen_deals:
                continue
            seen_deals.add(deal_id)

            lead_count = to_int(row.get(lead_count_field)) if lead_count_field else 1
            row_stage_id = text(row.get(stage_field))
            comment_digits = normalize_phone(row.get(comment_field))
            skip_by_stage = bool(row_stage_id and row_stage_id in excluded_stage_ids)
            skip_by_phone = any(phone in comment_digits for phone in excluded_comment_phones)
            if skip_by_stage or skip_by_phone:
                excluded_count += lead_count
                if skip_by_stage and skip_by_phone:
                    excluded_stage_and_phone += lead_count
                elif skip_by_stage:
                    excluded_stage_only += lead_count
                else:
                    excluded_phone_only += lead_count
                continue

            day = parse_date(row.get(date_field))
            date_key = day.isoformat() if day else "unknown"
            stage_info = stage_map.get(row_stage_id, {})
            stage_semantics = text(stage_info.get("semantics")).upper()
            stage_name = text(stage_info.get("name") or row_stage_id or "Без стадии")
            stage_lookup = {normalize_key(row_stage_id), normalize_key(stage_name)}
            quality_key = "bad" if stage_semantics == "F" or bool(stage_lookup & untouched_stage_names) else "quality"
            add_leads(by_campaign, ALL_LEADS_KEY, date_key, lead_count)
            add_lead_quality(quality_by_campaign, ALL_LEADS_KEY, date_key, quality_key, lead_count)
            date_campaign_key = campaign_key_for_date(campaign_date_rules, day)
            if date_campaign_key:
                add_leads(by_campaign, date_campaign_key, date_key, lead_count)
                add_lead_quality(quality_by_campaign, date_campaign_key, date_key, quality_key, lead_count)
                date_mapped_count += lead_count
            if utm_campaign_key:
                add_leads(by_campaign, utm_campaign_key, date_key, lead_count)
                add_leads(by_utm_campaign, utm_campaign_key, date_key, lead_count)
                add_lead_quality(quality_by_campaign, utm_campaign_key, date_key, quality_key, lead_count)
                utm_campaign_count += lead_count
            else:
                unknown_utm_campaign_count += lead_count
            add_stage_summary(stage_summary, row_stage_id, stage_name, quality_key, date_key, lead_count)
            matched_count += lead_count

    status = (
        f"Сделки из воронки «{deal_category_name or deal_category_id}»: {matched_count} шт.; "
        f"отбор по utm_source + utm_medium; UTM-групп: {len(utm_groups)}"
    )
    if excluded_count:
        status += f"; исключено: {excluded_count} шт."
        reason_parts = []
        if excluded_stage_only:
            reason_parts.append(f"стадии: {excluded_stage_only}")
        if excluded_phone_only:
            reason_parts.append(f"телефон: {excluded_phone_only}")
        if excluded_stage_and_phone:
            reason_parts.append(f"стадия+телефон: {excluded_stage_and_phone}")
        status += " (" + ", ".join(reason_parts) + ")"
    if missing_stage_names:
        status += "; не найдены стадии: " + ", ".join(missing_stage_names)
    if stage_map_error:
        status += f"; качество стадий: не удалось получить справочник ({stage_map_error})"
    if matched_count:
        status += f"; utm_campaign для креативов: {utm_campaign_count} шт."
        if unknown_utm_campaign_count:
            status += f", без utm_campaign: {unknown_utm_campaign_count} шт."
    if campaign_date_rules:
        status += f"; по датам кампаний распределено: {date_mapped_count} шт."
    all_quality = quality_by_campaign.get(ALL_LEADS_KEY, {})
    quality_total = sum(to_int(item.get("quality")) for item in all_quality.values())
    bad_total = sum(to_int(item.get("bad")) for item in all_quality.values())
    if quality_total or bad_total:
        status += f"; качество: {quality_total} лидов, {bad_total} обращений"

    return LeadData(
        True,
        "Bitrix",
        status,
        by_campaign,
        quality_by_campaign,
        stage_summary,
        by_utm_campaign,
    )


def load_leads(date_from: date | None = None, date_to: date | None = None) -> LeadData:
    return (
        load_leads_from_bitrix(date_from, date_to)
        or load_leads_from_csv(date_from, date_to)
        or LeadData(False, "не подключено", "Лиды пока не подключены: ожидается Bitrix или leads.csv", {})
    )


def lead_lookup_keys(info: dict[str, str], extra_keys: Iterable[str] | None = None) -> list[str]:
    result: list[str] = []
    for value in [info.get("campaign_id"), info.get("campaign_name"), *(extra_keys or [])]:
        for key in (normalize_key(value), normalize_utm_value(value)):
            if key and key not in result:
                result.append(key)
    return result


def campaign_leads(
    info: dict[str, str],
    lead_data: LeadData,
    fallback_all: bool = False,
    extra_keys: Iterable[str] | None = None,
) -> dict[str, int]:
    keys = lead_lookup_keys(info, extra_keys)
    result: dict[str, int] = {}
    for key in keys:
        if key and key in lead_data.by_campaign:
            for day, count in lead_data.by_campaign[key].items():
                result[day] = result.get(day, 0) + count
    if not result and fallback_all:
        for day, count in lead_data.by_campaign.get(ALL_LEADS_KEY, {}).items():
            result[day] = result.get(day, 0) + count
    return result


def campaign_lead_quality(
    info: dict[str, str],
    lead_data: LeadData,
    fallback_all: bool = False,
    extra_keys: Iterable[str] | None = None,
) -> dict[str, dict[str, int]]:
    keys = lead_lookup_keys(info, extra_keys)
    return merge_lead_quality(lead_data.quality_by_campaign, keys, fallback_all)


def attach_leads(
    daily: list[dict[str, object]],
    total: dict[str, float],
    leads_by_day: dict[str, int],
    enabled: bool,
    quality_by_day: dict[str, dict[str, int]] | None = None,
    include_missing_lead_dates: bool = True,
) -> None:
    quality_by_day = quality_by_day or {}
    existing_dates = {text(row.get("date")) for row in daily}
    if include_missing_lead_dates:
        for day in sorted(set(leads_by_day) | set(quality_by_day)):
            if day and day not in existing_dates:
                daily.append(
                    {
                        "date": day,
                        "impressions": 0.0,
                        "clicks": 0.0,
                        "spend": 0.0,
                        "bonus_spend": 0.0,
                        "ctr": 0.0,
                        "cpc": 0.0,
                        "cpm": 0.0,
                    }
                )
    else:
        leads_by_day = {day: count for day, count in leads_by_day.items() if day in existing_dates}
        quality_by_day = {day: value for day, value in quality_by_day.items() if day in existing_dates}
    total_leads = sum(leads_by_day.values())
    daily.sort(key=lambda row: text(row.get("date")))
    for row in daily:
        day = text(row.get("date"))
        row["leads"] = leads_by_day.get(day, 0)
        quality = quality_by_day.get(day, {})
        row["quality_leads"] = to_int(quality.get("quality"))
        row["bad_leads"] = to_int(quality.get("bad"))
        if enabled and not quality_by_day:
            row["quality_leads"] = row["leads"]
        quality_base = to_float(row.get("quality_leads")) + to_float(row.get("bad_leads"))
        row["lead_quality_rate"] = to_float(row.get("quality_leads")) / quality_base * 100 if enabled and quality_base else None
        row["click_to_lead"] = qualified_leads(row) / to_float(row.get("clicks")) * 100 if enabled and to_float(row.get("clicks")) else None
    total["leads"] = float(total_leads)
    total["quality_leads"] = float(sum(to_int(item.get("quality")) for item in quality_by_day.values()))
    total["bad_leads"] = float(sum(to_int(item.get("bad")) for item in quality_by_day.values()))
    if enabled and not quality_by_day:
        total["quality_leads"] = float(total_leads)
    total["click_to_lead"] = qualified_leads(total) / total["clicks"] * 100 if enabled and total["clicks"] else None
    total["impression_to_lead"] = qualified_leads(total) / total["impressions"] * 100 if enabled and total["impressions"] else None
    quality_total = to_float(total.get("quality_leads")) + to_float(total.get("bad_leads"))
    total["lead_quality_rate"] = to_float(total.get("quality_leads")) / quality_total * 100 if enabled and quality_total else None


def campaign_summary(export_group: dict[str, Export], lead_data: LeadData, owns_bitrix_avito_leads: bool = False) -> dict[str, object] | None:
    performance = export_group.get("performance")
    if not performance:
        return None
    info = performance.info or next((export.info for export in export_group.values() if export.info), {})
    campaign_sheet = next(
        (
            name
            for name, rows in performance.sheets.items()
            if name.startswith("Кампания ") and rows and text(rows[0][0]) == "Дата"
        ),
        "",
    )
    if not campaign_sheet:
        return None

    daily, total = parse_metric_sheet(performance.sheets[campaign_sheet])
    if lead_data.source == "Bitrix":
        leads_by_day = lead_data.by_campaign.get(ALL_LEADS_KEY, {}) if owns_bitrix_avito_leads else {}
        quality_by_day = lead_data.quality_by_campaign.get(ALL_LEADS_KEY, {}) if owns_bitrix_avito_leads else {}
    else:
        leads_by_day = campaign_leads(info, lead_data)
        quality_by_day = campaign_lead_quality(info, lead_data)
    attach_leads(daily, total, leads_by_day, lead_data.enabled, quality_by_day)

    groups = [
        summarize_entity(name, "Группа", rows)
        for name, rows in performance.sheets.items()
        if name.startswith("Группа ") and rows and text(rows[0][0]) == "Дата"
    ]
    creatives = [
        summarize_entity(name, "Креатив", rows)
        for name, rows in performance.sheets.items()
        if name.startswith("Креатив ") and rows and text(rows[0][0]) == "Дата"
    ]

    demography = {"gender": [], "income": [], "age": {"rows": [], "totals": [], "by_gender": {}}}
    if export_group.get("demography"):
        sheets = export_group["demography"].sheets
        for name, rows in sheets.items():
            if name.startswith("Кампания ") and name.endswith(" Пол"):
                demography["gender"] = parse_distribution(rows, "Пол")
            elif name.startswith("Кампания ") and name.endswith(" Доход"):
                demography["income"] = parse_distribution(rows, "Доход")
            elif name.startswith("Кампания ") and name.endswith(" Возраст"):
                demography["age"] = parse_age(rows)

    geography = {"campaign": []}
    if export_group.get("geography"):
        sheets = export_group["geography"].sheets
        for name, rows in sheets.items():
            if name.startswith("Кампания ") and name.endswith(" Гео"):
                geography["campaign"] = parse_distribution(rows, "Регион")

    total_clicks = total["clicks"]
    demography["gender"] = add_shares(demography["gender"], total_clicks)
    demography["income"] = add_shares(demography["income"], total_clicks)
    demography["age"]["totals"] = add_shares(demography["age"]["totals"], total_clicks)
    geography["campaign"] = add_shares(geography["campaign"], total_clicks)

    dates = [parse_date(row["date"]) for row in daily if parse_date(row["date"])]
    files = {kind: export.path.name for kind, export in export_group.items()}
    display_name = performance.display_name or info.get("display_name") or info.get("campaign_name", "")
    return {
        "id": info.get("campaign_id", ""),
        "name": display_name,
        "source_name": info.get("campaign_name", ""),
        "info": info,
        "files": files,
        "daily": daily,
        "total": total,
        "groups": sorted(groups, key=lambda item: to_float(item.get("clicks")), reverse=True),
        "creatives": sorted(creatives, key=lambda item: to_float(item.get("clicks")), reverse=True),
        "demography": demography,
        "geography": geography,
        "data_start": min(dates).isoformat() if dates else "",
        "data_end": max(dates).isoformat() if dates else "",
        "active_days": len(active_rows(daily)),
        "zero_gap": longest_zero_gap(daily),
    }


def merge_daily(campaigns: list[dict[str, object]], lead_enabled: bool) -> list[dict[str, object]]:
    by_day: dict[str, dict[str, object]] = {}
    for campaign in campaigns:
        for row in campaign["daily"]:
            day = text(row.get("date"))
            bucket = by_day.setdefault(
                day,
                {
                    "date": day,
                    "impressions": 0.0,
                    "clicks": 0.0,
                    "spend": 0.0,
                    "bonus_spend": 0.0,
                    "leads": 0.0,
                    "quality_leads": 0.0,
                    "bad_leads": 0.0,
                },
            )
            for key in ("impressions", "clicks", "spend", "bonus_spend", "leads", "quality_leads", "bad_leads"):
                bucket[key] = to_float(bucket.get(key)) + to_float(row.get(key))
    result = []
    for day in sorted(by_day):
        row = by_day[day]
        complete_rates(row)
        row["click_to_lead"] = qualified_leads(row) / to_float(row.get("clicks")) * 100 if lead_enabled and to_float(row.get("clicks")) else None
        quality_base = to_float(row.get("quality_leads")) + to_float(row.get("bad_leads"))
        row["lead_quality_rate"] = to_float(row.get("quality_leads")) / quality_base * 100 if lead_enabled and quality_base else None
        result.append(row)
    return result


def summarize_totals(campaigns: list[dict[str, object]], lead_enabled: bool) -> dict[str, float | None]:
    total = empty_total()
    lead_total = 0.0
    quality_total = 0.0
    bad_total = 0.0
    for campaign in campaigns:
        campaign_total = campaign["total"]
        for key in ("impressions", "clicks", "spend", "bonus_spend"):
            total[key] += to_float(campaign_total.get(key))
        lead_total += to_float(campaign_total.get("leads"))
        quality_total += to_float(campaign_total.get("quality_leads"))
        bad_total += to_float(campaign_total.get("bad_leads"))
    complete_rates(total)
    total["leads"] = lead_total
    total["quality_leads"] = quality_total
    total["bad_leads"] = bad_total
    total["click_to_lead"] = quality_total / total["clicks"] * 100 if lead_enabled and total["clicks"] else None
    total["impression_to_lead"] = quality_total / total["impressions"] * 100 if lead_enabled and total["impressions"] else None
    quality_base = quality_total + bad_total
    total["lead_quality_rate"] = quality_total / quality_base * 100 if lead_enabled and quality_base else None
    return total


def apply_overall_leads(daily: list[dict[str, object]], total: dict[str, float | None], lead_data: LeadData) -> None:
    if not lead_data.enabled:
        return
    leads_by_day = lead_data.by_campaign.get(ALL_LEADS_KEY, {})
    quality_by_day = lead_data.quality_by_campaign.get(ALL_LEADS_KEY, {})
    if not leads_by_day:
        return

    for row in daily:
        day = text(row.get("date"))
        row["leads"] = leads_by_day.get(day, 0)
        quality = quality_by_day.get(day, {})
        row["quality_leads"] = to_int(quality.get("quality"))
        row["bad_leads"] = to_int(quality.get("bad"))
        row["click_to_lead"] = (
            qualified_leads(row) / to_float(row.get("clicks")) * 100
            if to_float(row.get("clicks"))
            else None
        )
        quality_base = to_float(row.get("quality_leads")) + to_float(row.get("bad_leads"))
        row["lead_quality_rate"] = to_float(row.get("quality_leads")) / quality_base * 100 if quality_base else None

    total_leads = float(sum(leads_by_day.values()))
    total_quality = float(sum(to_int(item.get("quality")) for item in quality_by_day.values()))
    total_bad = float(sum(to_int(item.get("bad")) for item in quality_by_day.values()))
    total["leads"] = total_leads
    total["click_to_lead"] = total_quality / to_float(total.get("clicks")) * 100 if to_float(total.get("clicks")) else None
    total["impression_to_lead"] = (
        total_quality / to_float(total.get("impressions")) * 100
        if to_float(total.get("impressions"))
        else None
    )
    total["quality_leads"] = total_quality
    total["bad_leads"] = total_bad
    quality_base = total_quality + total_bad
    total["lead_quality_rate"] = total_quality / quality_base * 100 if quality_base else None


def average(rows: list[dict[str, object]], key: str) -> float:
    if not rows:
        return 0.0
    return sum(to_float(row.get(key)) for row in rows) / len(rows)


def stddev(rows: list[dict[str, object]], key: str) -> float:
    values = [to_float(row.get(key)) for row in rows]
    if len(values) < 2:
        return 0.0
    return statistics.pstdev(values)


def forecast(rows: list[dict[str, object]], lead_enabled: bool) -> dict[str, object]:
    active = active_rows(rows)
    lead_key = "quality_leads" if lead_enabled else "leads"
    avg = {key: average(active, key) for key in ("impressions", "clicks", "spend")}
    avg["leads"] = average(active, lead_key)
    sd = {key: stddev(active, key) for key in ("clicks", "spend")}
    sd["leads"] = stddev(active, lead_key)
    projections = {}
    for horizon in (7, 30):
        item = {
            "impressions": avg["impressions"] * horizon,
            "clicks": avg["clicks"] * horizon,
            "spend": avg["spend"] * horizon,
            "clicks_low": max(0.0, (avg["clicks"] - sd["clicks"]) * horizon),
            "clicks_high": (avg["clicks"] + sd["clicks"]) * horizon,
            "spend_low": max(0.0, (avg["spend"] - sd["spend"]) * horizon),
            "spend_high": (avg["spend"] + sd["spend"]) * horizon,
        }
        if lead_enabled:
            item["leads"] = avg["leads"] * horizon
            item["leads_low"] = max(0.0, (avg["leads"] - sd["leads"]) * horizon)
            item["leads_high"] = (avg["leads"] + sd["leads"]) * horizon
        projections[str(horizon)] = item
    confidence = "низкая" if len(active) < 7 else "средняя" if len(active) < 14 else "нормальная"
    return {"active_days": len(active), "avg": avg, "projections": projections, "confidence": confidence}


def report_period(campaign_exports: dict[str, dict[str, Export]]) -> tuple[date | None, date | None]:
    starts: list[date] = []
    ends: list[date] = []
    for exports in campaign_exports.values():
        for export in exports.values():
            start = parse_date(export.info.get("period_start"))
            end = parse_date(export.info.get("period_end"))
            if start:
                starts.append(start)
            if end:
                ends.append(end)
    return (min(starts) if starts else None, max(ends) if ends else None)


def api_data_files() -> list[Path]:
    if not AVITO_DATA_DIR.exists():
        return []
    return sorted(AVITO_DATA_DIR.glob("*.json"))


def normalize_metric_row(row: dict[str, object]) -> dict[str, object]:
    metrics = {
        "impressions": to_float(row.get("impressions") or row.get("shows") or row.get("views")),
        "clicks": to_float(row.get("clicks")),
        "spend": to_float(row.get("spend") or row.get("cost") or row.get("expense")),
        "bonus_spend": to_float(row.get("bonus_spend") or row.get("bonusSpend") or row.get("bonus")),
    }
    return complete_rates(metrics)


def campaign_from_api_data(
    payload: dict[str, object],
    lead_data: LeadData,
    owns_all_leads: bool = False,
) -> dict[str, object]:
    campaign = payload.get("campaign") if isinstance(payload.get("campaign"), dict) else {}
    campaign_id = text(campaign.get("id") or payload.get("campaign_id") or "druzheskiy")
    campaign_name = text(campaign.get("name") or payload.get("campaign_name") or REPORT_CAMPAIGN_NAME)
    lead_keys = as_list(campaign.get("lead_keys") or campaign.get("utm_campaigns") or campaign.get("utm_campaign"))
    lead_date_from = parse_date(campaign.get("lead_date_from"))
    lead_date_to = parse_date(campaign.get("lead_date_to"))
    campaign_info = {"campaign_id": campaign_id, "campaign_name": campaign_name}
    daily = []
    for row in payload.get("daily", []) if isinstance(payload.get("daily"), list) else []:
        if not isinstance(row, dict):
            continue
        day = parse_date(row.get("date") or row.get("day"))
        if not day:
            continue
        daily.append({"date": day.isoformat(), **normalize_metric_row(row)})
    daily.sort(key=lambda row: text(row.get("date")))
    total = sum_metric_rows(daily)
    lead_lookup_info = {"campaign_id": "", "campaign_name": ""} if lead_keys else campaign_info
    leads_by_day = campaign_leads(lead_lookup_info, lead_data, fallback_all=owns_all_leads, extra_keys=lead_keys)
    quality_by_day = campaign_lead_quality(lead_lookup_info, lead_data, fallback_all=owns_all_leads, extra_keys=lead_keys)
    if lead_date_from or lead_date_to:
        leads_by_day = {
            day: count
            for day, count in leads_by_day.items()
            if in_date_range(parse_date(day), lead_date_from, lead_date_to)
        }
        quality_by_day = {
            day: value
            for day, value in quality_by_day.items()
            if in_date_range(parse_date(day), lead_date_from, lead_date_to)
        }
    attach_leads(
        daily,
        total,
        leads_by_day,
        lead_data.enabled,
        quality_by_day,
        include_missing_lead_dates=bool(lead_date_from or lead_date_to),
    )

    def normalize_entities(items: object) -> list[dict[str, object]]:
        result = []
        for item in items if isinstance(items, list) else []:
            if not isinstance(item, dict):
                continue
            label = text(item.get("id") or item.get("name") or item.get("label"))
            entity_daily = []
            for row in item.get("daily", []) if isinstance(item.get("daily"), list) else []:
                if not isinstance(row, dict):
                    continue
                day = parse_date(row.get("date") or row.get("day"))
                if not day:
                    continue
                entity_daily.append({"date": day.isoformat(), **normalize_metric_row(row)})
            entity_daily.sort(key=lambda row: text(row.get("date")))
            metrics = sum_metric_rows(entity_daily) if entity_daily else normalize_metric_row(item)
            result.append({"id": text(item.get("id") or label), "label": text(item.get("label") or item.get("name") or label), "daily": entity_daily, **metrics})
        return sorted(result, key=lambda item: to_float(item.get("clicks")), reverse=True)

    total_clicks = total["clicks"]
    demography = payload.get("demography") if isinstance(payload.get("demography"), dict) else {}
    geography = payload.get("geography") if isinstance(payload.get("geography"), dict) else {}
    dates = [parse_date(row["date"]) for row in daily if parse_date(row["date"])]
    return {
        "id": campaign_id,
        "name": campaign_name,
        "source_name": text(campaign.get("source_name") or ""),
        "info": campaign_info,
        "files": {"api": text(payload.get("source") or "Avito API")},
        "daily": daily,
        "total": total,
        "groups": normalize_entities(payload.get("groups")),
        "creatives": normalize_entities(payload.get("creatives")),
        "demography": {
            "gender": add_shares(demography.get("gender", []) if isinstance(demography.get("gender"), list) else [], total_clicks),
            "income": add_shares(demography.get("income", []) if isinstance(demography.get("income"), list) else [], total_clicks),
            "age": {"rows": [], "totals": add_shares(demography.get("age", []) if isinstance(demography.get("age"), list) else [], total_clicks), "by_gender": {}},
        },
        "geography": {
            "campaign": add_shares(geography.get("campaign", []) if isinstance(geography.get("campaign"), list) else [], total_clicks)
        },
        "data_start": min(dates).isoformat() if dates else "",
        "data_end": max(dates).isoformat() if dates else "",
        "active_days": len(active_rows(daily)),
        "zero_gap": longest_zero_gap(daily),
    }


def placeholder_campaign(lead_data: LeadData, date_from: date | None, date_to: date | None) -> dict[str, object]:
    leads_by_day = lead_data.by_campaign.get(ALL_LEADS_KEY, {})
    quality_by_day = lead_data.quality_by_campaign.get(ALL_LEADS_KEY, {})
    daily = []
    for day in sorted(leads_by_day):
        daily.append({
            "date": day,
            "impressions": 0.0,
            "clicks": 0.0,
            "spend": 0.0,
            "bonus_spend": 0.0,
            "ctr": 0.0,
            "cpc": 0.0,
            "cpm": 0.0,
        })
    total = sum_metric_rows(daily)
    attach_leads(daily, total, leads_by_day, lead_data.enabled, quality_by_day)
    return {
        "id": REPORT_CAMPAIGN_SLUG,
        "name": REPORT_CAMPAIGN_NAME,
        "source_name": "",
        "info": {"campaign_id": REPORT_CAMPAIGN_SLUG, "campaign_name": REPORT_CAMPAIGN_NAME},
        "files": {},
        "daily": daily,
        "total": total,
        "groups": [],
        "creatives": [],
        "demography": {"gender": [], "income": [], "age": {"rows": [], "totals": [], "by_gender": {}}},
        "geography": {"campaign": []},
        "data_start": date_from.isoformat() if date_from else "",
        "data_end": date_to.isoformat() if date_to else "",
        "active_days": len(active_rows(daily)),
        "zero_gap": longest_zero_gap(daily),
    }


def load_api_campaigns(lead_data: LeadData) -> list[dict[str, object]]:
    payloads = []
    registry = campaign_registry_meta()
    for path in api_data_files():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            campaign = payload.get("campaign") if isinstance(payload.get("campaign"), dict) else {}
            campaign_id = text(campaign.get("id") or payload.get("campaign_id"))
            if campaign_id in registry:
                payload = {**payload, "campaign": {**registry[campaign_id], **campaign}}
            payloads.append(payload)
    owns_all_leads = len(payloads) == 1
    return [campaign_from_api_data(payload, lead_data, owns_all_leads=owns_all_leads) for payload in payloads]


def env_report_period() -> tuple[date | None, date | None]:
    env = read_env_file(BITRIX_ENV_FILE)
    start = parse_date(env.get("REPORT_START_DATE")) if env else None
    end = date.today()
    return start, end


def build_report() -> dict[str, object]:
    date_from, date_to = env_report_period()
    lead_data = load_leads(date_from, date_to)
    campaigns = load_api_campaigns(lead_data)
    if not campaigns:
        campaigns = [placeholder_campaign(lead_data, date_from, date_to)]
    campaigns.sort(key=lambda item: to_float(item["total"].get("clicks")), reverse=True)
    aggregate_daily = merge_daily(campaigns, lead_data.enabled)
    totals = summarize_totals(campaigns, lead_data.enabled)
    apply_overall_leads(aggregate_daily, totals, lead_data)
    dates = [parse_date(row["date"]) for row in aggregate_daily if parse_date(row["date"])]
    return {
        "lead_data": lead_data,
        "campaigns": campaigns,
        "daily": aggregate_daily,
        "total": totals,
        "forecast": forecast(aggregate_daily, lead_data.enabled),
        "data_start": min(dates).isoformat() if dates else (date_from.isoformat() if date_from else ""),
        "data_end": max(dates).isoformat() if dates else (date_to.isoformat() if date_to else ""),
        "active_days": len(active_rows(aggregate_daily)),
        "zero_gap": longest_zero_gap(aggregate_daily),
    }


def clip_label(value: object, limit: int = 24) -> str:
    label = text(value)
    return label if len(label) <= limit else label[: limit - 1] + "…"


def chart_tooltip(title: object, lines: Iterable[object]) -> str:
    body = "".join(f"<span>{esc(line)}</span>" for line in lines if text(line))
    return f"<strong>{esc(title)}</strong>{body}"


def svg_daily_combo(rows: list[dict[str, object]], height: int = 285) -> str:
    width = 980
    left, right, top, bottom = 58, 62, 34, 52
    chart_w = width - left - right
    chart_h = height - top - bottom
    data = rows
    if not data:
        return f'<svg viewBox="0 0 {width} {height}"><text x="24" y="56" class="svg-muted">Нет данных</text></svg>'

    labels = [text(row.get("date")) for row in data]
    clicks = [to_float(row.get("clicks")) for row in data]
    ctrs = [to_float(row.get("ctr")) for row in data]
    max_click_value = max(clicks) if clicks else 0.0
    max_ctr_value = max(ctrs) if ctrs else 0.0
    max_clicks = max(max_click_value * 1.15, 1)
    max_ctr = max(max_ctr_value * 1.25, 0.1) if max_ctr_value else 1
    slot = chart_w / max(1, len(data))
    bar_w = min(28, max(8, slot * 0.58))

    def x_center(pos: int) -> float:
        return left + slot * pos + slot / 2

    def y_click(value: float) -> float:
        return top + chart_h - chart_h * value / max_clicks

    def y_ctr(value: float) -> float:
        return top + chart_h - chart_h * value / max_ctr

    grid = []
    for step in range(4):
        value = max_clicks * step / 3
        yy = y_click(value)
        grid.append(
            f'<line x1="{left}" x2="{width-right}" y1="{yy:.1f}" y2="{yy:.1f}" class="grid-line" />'
            f'<text x="{left-9}" y="{yy+4:.1f}" text-anchor="end" class="axis-label">{esc(number(value))}</text>'
        )
    grid.append(f'<text x="{left}" y="18" class="axis-title">Клики</text>')
    grid.append(f'<text x="{width-right}" y="18" text-anchor="end" class="axis-title">CTR</text>')
    grid.append(f'<text x="{width-right+8}" y="{top+4}" class="axis-label">{esc(pct(max_ctr, 2))}</text>')
    grid.append(f'<text x="{width-right+8}" y="{top+chart_h+4}" class="axis-label">0%</text>')

    bars = []
    hovers = []
    for idx, value in enumerate(clicks):
        h = chart_h * value / max_clicks if max_clicks else 0
        x = x_center(idx) - bar_w / 2
        y = top + chart_h - h
        row = data[idx]
        tooltip = chart_tooltip(
            display_date(labels[idx]),
            [
                f"Столбец: клики за день — {number(value)}",
                f"Линия: CTR — {pct(row.get('ctr'), 2)}",
                f"Показы — {number(row.get('impressions'))}",
                f"Лиды — {number(qualified_leads(row))}",
                f"Обращения — {number(appeal_count(row))}",
            ],
        )
        bars.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{h:.1f}" rx="4" class="bar-clicks" data-chart-tooltip="{esc(tooltip)}" />'
        )
        hovers.append(
            f'<rect x="{left + slot * idx:.1f}" y="{top}" width="{slot:.1f}" height="{chart_h}" class="chart-hover" data-chart-tooltip="{esc(tooltip)}" />'
        )

    line_points = " ".join(f"{x_center(idx):.1f},{y_ctr(value):.1f}" for idx, value in enumerate(ctrs))
    dot_items = []
    for idx, value in enumerate(ctrs):
        if value <= 0:
            continue
        tooltip = chart_tooltip(
            display_date(labels[idx]),
            [
                f"CTR — {pct(value, 2)}",
                f"Клики — {number(clicks[idx])}",
                f"Показы — {number(data[idx].get('impressions'))}",
            ],
        )
        dot_items.append(
            f'<circle cx="{x_center(idx):.1f}" cy="{y_ctr(value):.1f}" r="3" class="dot-ctr" data-chart-tooltip="{esc(tooltip)}" />'
        )
    dots = "".join(dot_items)
    label_indexes = sorted({0, len(labels) // 2, len(labels) - 1})
    x_labels = "".join(
        f'<text x="{x_center(idx):.1f}" y="{height-16}" text-anchor="middle" class="axis-label">{esc(short_date(labels[idx]))}</text>'
        for idx in label_indexes
    )
    legend = """
      <g class="chart-legend" transform="translate(58 268)">
        <rect width="10" height="10" rx="2" class="bar-clicks"></rect><text x="16" y="10">Клики</text>
        <circle cx="82" cy="5" r="4" class="dot-ctr"></circle><text x="92" y="10">CTR</text>
      </g>
    """
    return f"""
    <svg viewBox="0 0 {width} {height}" role="img" aria-label="Клики и CTR по дням">
      {''.join(grid)}
      {''.join(bars)}
      <polyline points="{line_points}" fill="none" class="line-ctr" />
      {dots}
      {''.join(hovers)}
      {x_labels}
      {legend}
    </svg>
    """


def svg_bar(items: list[dict[str, object]], label_key: str, value_key: str, color_class: str, limit: int = 8) -> str:
    data = [item for item in items if to_float(item.get(value_key)) > 0][:limit]
    width = 820
    row_h = 34
    left, right, top = 220, 86, 12
    height = max(90, top * 2 + row_h * len(data))
    chart_w = width - left - right
    max_value = max((to_float(item.get(value_key)) for item in data), default=1)
    rows = []
    for idx, item in enumerate(data):
        y = top + idx * row_h
        value = to_float(item.get(value_key))
        bar_w = chart_w * value / max_value if max_value else 0
        label = text(item.get(label_key))
        tooltip = chart_tooltip(label, [f"Клики — {number(value)}"])
        rows.append(
            f'<text x="{left-12}" y="{y+21}" text-anchor="end" class="bar-label" data-chart-tooltip="{esc(tooltip)}">{esc(clip_label(label))}</text>'
            f'<rect x="{left}" y="{y+7}" width="{bar_w:.1f}" height="14" rx="4" class="{color_class}" data-chart-tooltip="{esc(tooltip)}" />'
            f'<text x="{left+bar_w+8:.1f}" y="{y+21}" class="bar-value" data-chart-tooltip="{esc(tooltip)}">{esc(number(value))}</text>'
        )
    if not rows:
        rows.append('<text x="20" y="48" class="svg-muted">Нет данных</text>')
    return f'<svg viewBox="0 0 {width} {height}" role="img">{"".join(rows)}</svg>'


def funnel(total: dict[str, float | None], lead_enabled: bool) -> str:
    stages = [
        ("Показы", to_float(total.get("impressions")), "100%"),
        ("Клики", to_float(total.get("clicks")), pct(total.get("ctr"), 2)),
        ("Лиды", qualified_leads(total), pct(total.get("click_to_lead"), 2) if lead_enabled else "нет данных"),
    ]
    max_value = max(max((stage[1] for stage in stages), default=0), 1)
    rows = []
    for label, value, rate in stages:
        width = max(1.0, value / max_value * 100)
        muted = " muted" if label == "Лиды" and not lead_enabled else ""
        rows.append(
            f'<div class="funnel-row{muted}">'
            f'<div class="funnel-name">{esc(label)}</div>'
            f'<div class="funnel-track"><span style="width:{width:.2f}%"></span></div>'
            f'<div class="funnel-value"><strong>{esc(number(value)) if label != "Лиды" or lead_enabled else "нет данных"}</strong><small>{esc(rate)}</small></div>'
            f'</div>'
        )
    return "".join(rows)


def lead_stage_rows(stage_summary: dict[str, dict[str, object]]) -> list[dict[str, object]]:
    rows = []
    for item in stage_summary.values():
        if not isinstance(item, dict):
            continue
        count = to_int(item.get("count"))
        if count <= 0:
            continue
        rows.append(
            {
                "id": text(item.get("id")),
                "name": text(item.get("name")),
                "quality": text(item.get("quality")),
                "count": count,
                "by_day": item.get("by_day") if isinstance(item.get("by_day"), dict) else {},
            }
        )
    return sorted(rows, key=lambda item: to_int(item.get("count")), reverse=True)


def lead_utm_campaign_rows(by_utm_campaign: dict[str, dict[str, int]]) -> list[dict[str, object]]:
    rows = []
    for campaign, by_day in by_utm_campaign.items():
        if campaign == ALL_LEADS_KEY or not text(campaign):
            continue
        count = sum(to_int(value) for value in by_day.values())
        if count <= 0:
            continue
        rows.append(
            {
                "utm_campaign": campaign,
                "count": count,
                "daily": [{"date": day, "count": to_float(value)} for day, value in sorted(by_day.items())],
            }
        )
    return sorted(rows, key=lambda item: to_int(item.get("count")), reverse=True)


def lead_quality(
    total: dict[str, float | None],
    lead_enabled: bool,
    stage_summary: dict[str, dict[str, object]] | None = None,
) -> str:
    if not lead_enabled:
        return '<div class="quality-empty">нет данных</div>'
    quality = to_float(total.get("quality_leads"))
    bad = to_float(total.get("bad_leads"))
    base = quality + bad
    if not base:
        return '<div class="quality-empty">Лидов за период нет</div>'
    rows = [
        ("Лиды", quality, quality / base * 100, "quality-good"),
        ("Обращения", bad, bad / base * 100, "quality-bad"),
    ]
    summary = "".join(
        f'<div class="quality-row {esc(cls)}">'
        f'<span>{esc(label)}</span>'
        f'<strong>{esc(number(value))}</strong>'
        f'<small>{esc(pct(rate, 1))}</small>'
        f'</div>'
        for label, value, rate, cls in rows
    )
    stage_rows = lead_stage_rows(stage_summary or {})
    if stage_rows:
        stage_html = "".join(
            f'<div class="quality-stage {esc("quality-bad" if item["quality"] == "bad" else "quality-good")}">'
            f'<span>{esc(item["name"])}</span>'
            f'<strong>{esc(number(item["count"]))}</strong>'
            f'<small>{esc("обращение" if item["quality"] == "bad" else "лид")}</small>'
            f'</div>'
            for item in stage_rows[:6]
        )
        summary += f'<div class="quality-stages"><div class="quality-stage-title">Стадии</div>{stage_html}</div>'
    return summary


def table(headers: list[str], rows: list[list[object]], numeric: set[int] | None = None) -> str:
    numeric = numeric or set()
    head = "".join(f"<th>{esc(header)}</th>" for header in headers)
    body = []
    for row in rows:
        cells = []
        for idx, value in enumerate(row):
            cls = ' class="num"' if idx in numeric else ""
            cells.append(f"<td{cls}>{esc(value)}</td>")
        body.append(f"<tr>{''.join(cells)}</tr>")
    if not body:
        body.append(f'<tr class="empty-row"><td colspan="{len(headers)}">Нет данных</td></tr>')
    return f'<div class="table-wrap"><table class="data-table"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def campaign_table(campaigns: list[dict[str, object]], lead_enabled: bool) -> str:
    rows = []
    for campaign in campaigns:
        total = campaign["total"]
        rows.append(
            [
                campaign["name"] or campaign["id"],
                campaign["id"],
                number(total["impressions"]),
                number(total["clicks"]),
                pct(total["ctr"], 2),
                number(qualified_leads(total)) if lead_enabled else "нет данных",
                number(appeal_count(total)) if lead_enabled else "нет данных",
                pct(total["click_to_lead"], 2) if lead_enabled else "нет данных",
                money(total["spend"]),
                money(total["cpc"]),
            ]
        )
    return table(
        ["Кампания", "ID", "Показы", "Клики", "Показ → клик", "Лиды", "Обращения", "Клик → лид", "Расход", "CPC"],
        rows,
        numeric={2, 3, 4, 5, 6, 7, 8, 9},
    )


def entity_table(items: list[dict[str, object]], label: str) -> str:
    rows = [
        [
            item.get("id") or item.get("label") or item.get("index") or "",
            number(item.get("impressions")),
            number(item.get("clicks")),
            pct(item.get("ctr"), 2),
            money(item.get("spend")),
            money(item.get("cpc")),
        ]
        for item in items
        if to_float(item.get("impressions")) > 0 or to_float(item.get("clicks")) > 0
    ]
    return table([label, "Показы", "Клики", "Показ → клик", "Расход", "CPC"], rows, numeric={1, 2, 3, 4, 5})


def distribution_table(items: list[dict[str, object]], label: str) -> str:
    return table(
        [label, "Клики", "Доля"],
        [[item.get("label"), number(item.get("clicks")), pct(item.get("share"), 1)] for item in items if to_float(item.get("clicks")) > 0],
        numeric={1, 2},
    )


def kpi(title: str, value: str, note: str = "", accent: str = "accent-cyan", key: str = "") -> str:
    key_attr = f' data-kpi="{esc(key)}"' if key else ""
    value_id = f' id="kpi-{esc(key)}"' if key else ""
    note_id = f' id="kpi-{esc(key)}-note"' if key else ""
    return f'<article class="kpi-card {esc(accent)}"{key_attr}><span class="kpi-label">{esc(title)}</span><strong class="kpi-value"{value_id}>{esc(value)}</strong><small{note_id}>{esc(note)}</small></article>'


def overview_stat(title: str, value: str, key: str = "") -> str:
    key_attr = f' data-overview="{esc(key)}"' if key else ""
    value_id = f' id="overview-{esc(key)}"' if key else ""
    return f'<article class="overview-stat"{key_attr}><span>{esc(title)}</span><strong{value_id}>{esc(value)}</strong></article>'


def overview_period_stat(start: object, end: object) -> str:
    start_text = display_date(start)
    end_text = display_date(end)
    if start_text and end_text:
        value = f'<span id="overview-period-start">{esc(start_text)}</span><span class="period-separator">—</span><span id="overview-period-end">{esc(end_text)}</span>'
    else:
        value = '<span id="overview-period-start">нет данных</span><span class="period-separator"></span><span id="overview-period-end"></span>'
    return f'<article class="overview-stat overview-stat--period" data-overview="period"><span>Период</span><strong id="overview-period" class="period-value">{value}</strong></article>'


def top_entities(campaigns: list[dict[str, object]], key: str, limit: int = 8) -> list[dict[str, object]]:
    items = []
    for campaign in campaigns:
        for item in campaign.get(key, []):
            label = item.get("id") or item.get("label") or item.get("index") or ""
            items.append({**item, "label": f"{campaign['id']} · {label}"})
    return sorted(items, key=lambda item: to_float(item.get("clicks")), reverse=True)[:limit]


def clean_daily(rows: object) -> list[dict[str, object]]:
    result = []
    for row in rows if isinstance(rows, list) else []:
        if not isinstance(row, dict):
            continue
        result.append(
            {
                "date": text(row.get("date")),
                "impressions": to_float(row.get("impressions")),
                "clicks": to_float(row.get("clicks")),
                "spend": to_float(row.get("spend")),
                "bonus_spend": to_float(row.get("bonus_spend")),
                "leads": to_float(row.get("leads")),
                "quality_leads": to_float(row.get("quality_leads")),
                "bad_leads": to_float(row.get("bad_leads")),
            }
        )
    return result


def clean_entities(items: object) -> list[dict[str, object]]:
    result = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        result.append(
            {
                "id": text(item.get("id") or item.get("label")),
                "label": text(item.get("label") or item.get("id")),
                "daily": clean_daily(item.get("daily")),
                "impressions": to_float(item.get("impressions")),
                "clicks": to_float(item.get("clicks")),
                "spend": to_float(item.get("spend")),
                "bonus_spend": to_float(item.get("bonus_spend")),
            }
        )
    return result


def client_report_payload(report: dict[str, object]) -> dict[str, object]:
    return {
        "leadEnabled": bool(report["lead_data"].enabled),
        "dataStart": text(report.get("data_start")),
        "dataEnd": text(report.get("data_end")),
        "leadStages": [
            {
                "id": text(item.get("id")),
                "name": text(item.get("name")),
                "quality": text(item.get("quality")),
                "daily": [
                    {"date": text(day), "count": to_float(count)}
                    for day, count in (item.get("by_day") or {}).items()
                ],
            }
            for item in lead_stage_rows(report["lead_data"].stage_summary)
        ],
        "leadUtmCampaigns": lead_utm_campaign_rows(report["lead_data"].by_utm_campaign),
        "campaigns": [
            {
                "id": text(campaign.get("id")),
                "name": text(campaign.get("name") or campaign.get("id")),
                "daily": clean_daily(campaign.get("daily")),
                "groups": clean_entities(campaign.get("groups")),
                "creatives": clean_entities(campaign.get("creatives")),
            }
            for campaign in report.get("campaigns", [])
            if isinstance(campaign, dict)
        ],
    }


def render_campaign(campaign: dict[str, object], lead_enabled: bool, total_count: int) -> str:
    total = campaign["total"]
    open_attr = " open" if total_count <= 2 else ""
    groups = campaign["groups"]
    creatives = campaign["creatives"]
    demo = campaign["demography"]
    geo = campaign["geography"]["campaign"]

    return f"""
    <details class="campaign"{open_attr}>
      <summary>
        <span class="campaign-title">{esc(campaign["name"] or "Кампания")} <em>{esc(campaign["id"])}</em></span>
        <span>{esc(number(total["clicks"]))} кликов · {esc(pct(total["ctr"], 2))} · {esc(money(total["spend"]))}</span>
      </summary>
      <div class="campaign-body">
        <div class="mini-kpis">
          {kpi("Показы", number(total["impressions"]), accent="accent-cyan")}
          {kpi("Клики", number(total["clicks"]), accent="accent-coral")}
          {kpi("Показ → клик", pct(total["ctr"], 2), accent="accent-violet")}
          {kpi("Лиды", number(qualified_leads(total)) if lead_enabled else "нет данных", accent="accent-lime")}
          {kpi("Обращения", number(appeal_count(total)) if lead_enabled else "нет данных", accent="accent-coral")}
          {kpi("Клик → лид", pct(total["click_to_lead"], 2) if lead_enabled else "нет данных", accent="accent-amber")}
          {kpi("Расход", money(total["spend"]), accent="accent-emerald")}
        </div>
        <section class="panel panel-inner">
          <div class="panel-head"><h3>Динамика</h3><p>Столбцы — клики, линия — CTR</p></div>
          {svg_daily_combo(campaign["daily"], 250)}
        </section>
        <div class="table-grid">
          <article class="panel">
            <div class="panel-head"><h3>Группы</h3></div>
            {entity_table(groups, "Группа")}
          </article>
          <article class="panel">
            <div class="panel-head"><h3>Креативы</h3></div>
            {entity_table(creatives, "Креатив")}
          </article>
        </div>
        <div class="table-grid">
          <article class="panel">
            <div class="panel-head"><h3>Аудитория</h3></div>
            {distribution_table(demo["gender"], "Пол")}
            {distribution_table(demo["income"], "Доход")}
          </article>
          <article class="panel">
            <div class="panel-head"><h3>География</h3></div>
            {distribution_table(geo, "Регион")}
          </article>
        </div>
      </div>
    </details>
    """


def render_html(report: dict[str, object]) -> str:
    campaigns = report["campaigns"]
    lead_data = report["lead_data"]
    total = report["total"]
    daily = report["daily"]
    lead_enabled = lead_data.enabled
    projection_7 = report["forecast"]["projections"]["7"]
    projection_30 = report["forecast"]["projections"]["30"]
    top_groups = top_entities(campaigns, "groups", 7)
    top_creatives = top_entities(campaigns, "creatives", 7)
    report_data_json = json.dumps(client_report_payload(report), ensure_ascii=False).replace("</", "<\\/")

    if not campaigns:
        main_content = '<section class="panel panel-wide"><div class="panel-head"><h2>Нет данных</h2></div></section>'
    else:
        main_content = f"""
        <section id="summary-kpis" class="kpi-grid" aria-label="Ключевые показатели">
          {kpi("Показы", number(total["impressions"]), accent="accent-cyan", key="impressions")}
          {kpi("Клики", number(total["clicks"]), accent="accent-coral", key="clicks")}
          {kpi("Показ → клик", pct(total["ctr"], 2), accent="accent-violet", key="ctr")}
          {kpi("Лиды", number(qualified_leads(total)) if lead_enabled else "нет данных", accent="accent-lime", key="leads")}
          {kpi("Обращения", number(appeal_count(total)) if lead_enabled else "нет данных", accent="accent-coral", key="appeals")}
          {kpi("Клик → лид", pct(total["click_to_lead"], 2) if lead_enabled else "нет данных", accent="accent-amber", key="click-to-lead")}
          {kpi("Расход", money(total["spend"]), "CPC " + money(total["cpc"]), "accent-emerald", key="spend")}
        </section>

        <section class="filters-band" aria-label="Управление отчетом">
          <label class="field">
            <span>С даты</span>
            <input id="date-from" type="date" value="{esc(report["data_start"])}" min="{esc(report["data_start"])}" max="{esc(report["data_end"])}">
          </label>
          <label class="field">
            <span>По дату</span>
            <input id="date-to" type="date" value="{esc(report["data_end"])}" min="{esc(report["data_start"])}" max="{esc(report["data_end"])}">
          </label>
          <label class="field field-search">
            <span>Поиск</span>
            <input id="table-search" type="search" placeholder="Кампания, группа, креатив, UTM">
          </label>
          <button id="reset-period" class="ghost-button" type="button">Весь период</button>
        </section>

        <section class="chart-grid chart-grid--primary">
          <article class="panel panel-wide">
            <div class="panel-head"><h2>Динамика по дням</h2><p>Столбцы — клики, линия — CTR</p></div>
            <div id="daily-chart" class="chart-frame">{svg_daily_combo(daily)}</div>
          </article>
        </section>

        <section class="table-grid" id="campaigns">
          <article class="panel panel-wide">
            <div class="panel-head"><h2>Кампании</h2></div>
            <div id="campaign-table">{campaign_table(campaigns, lead_enabled)}</div>
          </article>
        </section>

        <section class="chart-grid chart-grid--secondary">
          <article class="panel">
            <div class="panel-head"><h2>Воронка</h2></div>
            <div id="funnel">{funnel(total, lead_enabled)}</div>
          </article>
          <article class="panel">
            <div class="panel-head"><h2>Качество лидов</h2><p>По стадиям Bitrix</p></div>
            <div id="lead-quality" class="quality-list">{lead_quality(total, lead_enabled, lead_data.stage_summary)}</div>
          </article>
          <article class="panel">
            <div class="panel-head"><h2>Прогноз</h2><p id="forecast-confidence">Надежность прогноза: {esc(report["forecast"]["confidence"])}</p></div>
            <div id="forecast-grid" class="forecast-grid">
              {kpi("7 дней", number(projection_7["clicks"]) + " кликов", money(projection_7["spend"]), "accent-coral")}
              {kpi("30 дней", number(projection_30["clicks"]) + " кликов", money(projection_30["spend"]), "accent-cyan")}
              {kpi("Лиды 30 дней", number(projection_30.get("leads", 0)) if lead_enabled else "нет данных", accent="accent-lime")}
            </div>
          </article>
          <article class="panel">
            <div class="panel-head"><h2>Группы</h2><p>Топ по кликам</p></div>
            <div id="groups-chart">{svg_bar(top_groups, "label", "clicks", "bar-blue", 7)}</div>
          </article>
          <article class="panel">
            <div class="panel-head"><h2>Креативы</h2><p>Топ по кликам</p></div>
            <div id="creatives-chart">{svg_bar(top_creatives, "label", "clicks", "bar-green", 7)}</div>
          </article>
        </section>

        <section id="details" class="details-section">
          <div class="panel-head panel-head-outside"><h2>Детализация</h2></div>
          {"".join(render_campaign(campaign, lead_enabled, len(campaigns)) for campaign in campaigns)}
        </section>
        """

    client_script = r"""
    const reportData = JSON.parse(document.getElementById('report-data').textContent);
    const searchInput = document.getElementById('table-search');
    const dateFromInput = document.getElementById('date-from');
    const dateToInput = document.getElementById('date-to');
    const navLinks = [...document.querySelectorAll('.site-nav a')];
    const sections = navLinks.map((link) => document.querySelector(link.getAttribute('href'))).filter(Boolean);
    const leadEnabled = Boolean(reportData.leadEnabled);
    const chartTooltip = document.getElementById('chart-tooltip');

    const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
    }[char]));
    const toNum = (value) => Number(value) || 0;
    const number = (value) => Math.round(toNum(value)).toLocaleString('ru-RU');
    const money = (value) => `${number(value)} ₽`;
    const decimal = (value, digits = 2) => toNum(value).toLocaleString('ru-RU', { minimumFractionDigits: digits, maximumFractionDigits: digits });
    const pct = (value, digits = 2) => value === null || value === undefined ? 'нет данных' : `${decimal(value, digits)}%`;
    const ruDate = (iso) => {
      if (!iso) return '';
      const [year, month, day] = iso.slice(0, 10).split('-');
      return year && month && day ? `${day}.${month}.${year}` : iso;
    };

    function hideChartTooltip() {
      if (!chartTooltip) return;
      chartTooltip.hidden = true;
      chartTooltip.classList.remove('is-visible');
    }

    function showChartTooltip(event, target) {
      if (!chartTooltip || !target?.dataset?.chartTooltip) return;
      chartTooltip.innerHTML = target.dataset.chartTooltip;
      chartTooltip.hidden = false;
      chartTooltip.classList.add('is-visible');
      const gap = 14;
      const margin = 12;
      const rect = chartTooltip.getBoundingClientRect();
      let left = event.clientX + gap;
      let top = event.clientY + gap;
      if (left + rect.width + margin > window.innerWidth) left = event.clientX - rect.width - gap;
      if (top + rect.height + margin > window.innerHeight) top = event.clientY - rect.height - gap;
      chartTooltip.style.left = `${Math.max(margin, left)}px`;
      chartTooltip.style.top = `${Math.max(margin, top)}px`;
    }

    document.addEventListener('pointermove', (event) => {
      const target = event.target instanceof Element ? event.target.closest('[data-chart-tooltip]') : null;
      if (!target) {
        hideChartTooltip();
        return;
      }
      showChartTooltip(event, target);
    });
    document.addEventListener('pointerleave', hideChartTooltip);

    function rowsInPeriod(rows, from, to) {
      return (rows || []).filter((row) => (!from || row.date >= from) && (!to || row.date <= to));
    }

    function complete(total) {
      total.quality_leads = toNum(total.quality_leads ?? total.qualityLeads);
      total.bad_leads = toNum(total.bad_leads ?? total.badLeads);
      total.ctr = total.impressions ? total.clicks / total.impressions * 100 : 0;
      total.cpc = total.clicks ? total.spend / total.clicks : 0;
      total.cpm = total.impressions ? total.spend / total.impressions * 1000 : 0;
      total.clickToLead = leadEnabled && total.clicks ? total.quality_leads / total.clicks * 100 : null;
      total.impressionToLead = leadEnabled && total.impressions ? total.quality_leads / total.impressions * 100 : null;
      const qualityBase = total.quality_leads + total.bad_leads;
      total.leadQualityRate = leadEnabled && qualityBase ? total.quality_leads / qualityBase * 100 : null;
      return total;
    }

    function appealCount(total) {
      return Math.max(0, toNum(total.leads) - toNum(total.quality_leads));
    }

    function sumRows(rows) {
      const total = rows.reduce((acc, row) => {
        acc.impressions += toNum(row.impressions);
        acc.clicks += toNum(row.clicks);
        acc.spend += toNum(row.spend);
        acc.bonus_spend += toNum(row.bonus_spend);
        acc.leads += toNum(row.leads);
        acc.quality_leads += toNum(row.quality_leads);
        acc.bad_leads += toNum(row.bad_leads);
        return acc;
      }, { impressions: 0, clicks: 0, spend: 0, bonus_spend: 0, leads: 0, quality_leads: 0, bad_leads: 0 });
      return complete(total);
    }

    function mergeDaily(campaigns, from, to) {
      const byDate = new Map();
      for (const campaign of campaigns) {
        for (const row of rowsInPeriod(campaign.daily, from, to)) {
          if (!byDate.has(row.date)) byDate.set(row.date, { date: row.date, impressions: 0, clicks: 0, spend: 0, bonus_spend: 0, leads: 0, quality_leads: 0, bad_leads: 0 });
          const target = byDate.get(row.date);
          target.impressions += toNum(row.impressions);
          target.clicks += toNum(row.clicks);
          target.spend += toNum(row.spend);
          target.bonus_spend += toNum(row.bonus_spend);
          target.leads += toNum(row.leads);
          target.quality_leads += toNum(row.quality_leads);
          target.bad_leads += toNum(row.bad_leads);
        }
      }
      return [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date)).map(complete);
    }

    function activeRows(rows) {
      return rows.filter((row) => toNum(row.impressions) > 0 || toNum(row.clicks) > 0 || toNum(row.spend) > 0);
    }

    function forecast(rows) {
      const active = activeRows(rows);
      const avg = (key) => active.length ? active.reduce((sum, row) => sum + toNum(row[key]), 0) / active.length : 0;
      const clicks = avg('clicks');
      const spend = avg('spend');
      const leads = leadEnabled ? avg('quality_leads') : avg('leads');
      return {
        confidence: active.length < 7 ? 'низкая' : active.length < 14 ? 'средняя' : 'нормальная',
        projections: {
          7: { clicks: clicks * 7, spend: spend * 7, leads: leads * 7 },
          30: { clicks: clicks * 30, spend: spend * 30, leads: leads * 30 },
        },
      };
    }

    function setText(id, value) {
      const node = document.getElementById(id);
      if (node) node.textContent = value;
    }

    function dailyTooltip(row) {
      return `<strong>${escapeHtml(ruDate(row.date) || row.date)}</strong><span>Столбец: клики за день — ${number(row.clicks)}</span><span>Линия: CTR — ${pct(row.ctr, 2)}</span><span>Показы — ${number(row.impressions)}</span><span>Лиды — ${number(row.quality_leads)}</span><span>Обращения — ${number(appealCount(row))}</span>`;
    }

    function svgDaily(rows) {
      const width = 980;
      const height = 285;
      const left = 58;
      const right = 62;
      const top = 34;
      const bottom = 52;
      const chartW = width - left - right;
      const chartH = height - top - bottom;
      if (!rows.length) return `<svg viewBox="0 0 ${width} ${height}"><text x="24" y="56" class="svg-muted">Нет данных</text></svg>`;
      const maxClickValue = Math.max(...rows.map((row) => toNum(row.clicks)), 0);
      const maxCtrValue = Math.max(...rows.map((row) => toNum(row.ctr)), 0);
      const maxClicks = Math.max(maxClickValue * 1.15, 1);
      const maxCtr = maxCtrValue ? Math.max(maxCtrValue * 1.25, 0.1) : 1;
      const slot = chartW / Math.max(1, rows.length);
      const barW = Math.min(28, Math.max(8, slot * 0.58));
      const xCenter = (idx) => left + slot * idx + slot / 2;
      const yClick = (value) => top + chartH - chartH * value / maxClicks;
      const yCtr = (value) => top + chartH - chartH * value / maxCtr;
      let grid = '';
      for (let step = 0; step < 4; step += 1) {
        const value = maxClicks * step / 3;
        const y = yClick(value);
        grid += `<line x1="${left}" x2="${width - right}" y1="${y.toFixed(1)}" y2="${y.toFixed(1)}" class="grid-line" />`;
        grid += `<text x="${left - 9}" y="${(y + 4).toFixed(1)}" text-anchor="end" class="axis-label">${number(value)}</text>`;
      }
      grid += `<text x="${left}" y="18" class="axis-title">Клики</text>`;
      grid += `<text x="${width - right}" y="18" text-anchor="end" class="axis-title">CTR</text>`;
      grid += `<text x="${width - right + 8}" y="${top + 4}" class="axis-label">${pct(maxCtr, 2)}</text>`;
      grid += `<text x="${width - right + 8}" y="${top + chartH + 4}" class="axis-label">0%</text>`;
      const bars = rows.map((row, idx) => {
        const value = toNum(row.clicks);
        const h = chartH * value / maxClicks;
        const x = xCenter(idx) - barW / 2;
        const y = top + chartH - h;
        return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${h.toFixed(1)}" rx="4" class="bar-clicks" data-chart-tooltip="${escapeHtml(dailyTooltip(row))}" />`;
      }).join('');
      const hovers = rows.map((row, idx) => `<rect x="${(left + slot * idx).toFixed(1)}" y="${top}" width="${slot.toFixed(1)}" height="${chartH}" class="chart-hover" data-chart-tooltip="${escapeHtml(dailyTooltip(row))}" />`).join('');
      const points = rows.map((row, idx) => `${xCenter(idx).toFixed(1)},${yCtr(toNum(row.ctr)).toFixed(1)}`).join(' ');
      const dots = rows.map((row, idx) => {
        if (toNum(row.ctr) <= 0) return '';
        return `<circle cx="${xCenter(idx).toFixed(1)}" cy="${yCtr(toNum(row.ctr)).toFixed(1)}" r="3" class="dot-ctr" data-chart-tooltip="${escapeHtml(dailyTooltip(row))}" />`;
      }).join('');
      const labelIndexes = [...new Set([0, Math.floor(rows.length / 2), rows.length - 1])].sort((a, b) => a - b);
      const labels = labelIndexes.map((idx) => `<text x="${xCenter(idx).toFixed(1)}" y="${height - 16}" text-anchor="middle" class="axis-label">${escapeHtml(ruDate(rows[idx].date).slice(0, 5))}</text>`).join('');
      const legend = `<g class="chart-legend" transform="translate(58 268)"><rect width="10" height="10" rx="2" class="bar-clicks"></rect><text x="16" y="10">Клики</text><circle cx="82" cy="5" r="4" class="dot-ctr"></circle><text x="92" y="10">CTR</text></g>`;
      return `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Клики и CTR по дням">${grid}${bars}<polyline points="${points}" fill="none" class="line-ctr" />${dots}${hovers}${labels}${legend}</svg>`;
    }

    function svgBar(items, colorClass, limit = 7) {
      const data = items.filter((item) => toNum(item.clicks) > 0).slice(0, limit);
      const width = 820;
      const rowH = 34;
      const left = 220;
      const right = 86;
      const top = 12;
      const height = Math.max(90, top * 2 + rowH * data.length);
      const chartW = width - left - right;
      const maxValue = Math.max(...data.map((item) => toNum(item.clicks)), 1);
      if (!data.length) return `<svg viewBox="0 0 ${width} ${height}" role="img"><text x="20" y="48" class="svg-muted">Нет данных</text></svg>`;
      const rows = data.map((item, idx) => {
        const y = top + idx * rowH;
        const value = toNum(item.clicks);
        const barW = chartW * value / maxValue;
        const label = String(item.label || item.id || '');
        const clipped = label.length <= 24 ? label : `${label.slice(0, 23)}…`;
        const tooltip = `<strong>${escapeHtml(label)}</strong><span>Клики — ${number(value)}</span>`;
        return `<text x="${left - 12}" y="${y + 21}" text-anchor="end" class="bar-label" data-chart-tooltip="${escapeHtml(tooltip)}">${escapeHtml(clipped)}</text><rect x="${left}" y="${y + 7}" width="${barW.toFixed(1)}" height="14" rx="4" class="${colorClass}" data-chart-tooltip="${escapeHtml(tooltip)}" /><text x="${(left + barW + 8).toFixed(1)}" y="${y + 21}" class="bar-value" data-chart-tooltip="${escapeHtml(tooltip)}">${number(value)}</text>`;
      }).join('');
      return `<svg viewBox="0 0 ${width} ${height}" role="img">${rows}</svg>`;
    }

    function entityTotals(campaigns, key, from, to) {
      const result = [];
      for (const campaign of campaigns) {
        for (const item of campaign[key] || []) {
          const total = item.daily && item.daily.length ? sumRows(rowsInPeriod(item.daily, from, to)) : complete({ impressions: toNum(item.impressions), clicks: toNum(item.clicks), spend: toNum(item.spend), bonus_spend: toNum(item.bonus_spend), leads: 0 });
          result.push({ ...total, id: item.id, label: `${campaign.id} · ${item.label || item.id}` });
        }
      }
      return result.sort((a, b) => toNum(b.clicks) - toNum(a.clicks));
    }

    function renderCampaignTable(campaigns) {
      const rows = campaigns.map((campaign) => {
        const total = campaign.total;
        return `<tr><td>${escapeHtml(campaign.name)}</td><td>${escapeHtml(campaign.id)}</td><td class="num">${number(total.impressions)}</td><td class="num">${number(total.clicks)}</td><td class="num">${pct(total.ctr, 2)}</td><td class="num">${leadEnabled ? number(total.quality_leads) : 'нет данных'}</td><td class="num">${leadEnabled ? number(appealCount(total)) : 'нет данных'}</td><td class="num">${leadEnabled ? pct(total.clickToLead, 2) : 'нет данных'}</td><td class="num">${money(total.spend)}</td><td class="num">${money(total.cpc)}</td></tr>`;
      }).join('') || '<tr class="empty-row"><td colspan="10">Нет данных</td></tr>';
      document.getElementById('campaign-table').innerHTML = `<div class="table-wrap"><table class="data-table"><thead><tr><th>Кампания</th><th>ID</th><th>Показы</th><th>Клики</th><th>Показ → клик</th><th>Лиды</th><th>Обращения</th><th>Клик → лид</th><th>Расход</th><th>CPC</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }

    function renderFunnel(total) {
      const maxValue = Math.max(total.impressions, total.clicks, total.quality_leads, 1);
      const stages = [
        ['Показы', total.impressions, '100%'],
        ['Клики', total.clicks, pct(total.ctr, 2)],
        ['Лиды', total.quality_leads, leadEnabled ? pct(total.clickToLead, 2) : 'нет данных'],
      ];
      document.getElementById('funnel').innerHTML = stages.map(([label, value, rate], idx) => {
        const width = Math.max(1, toNum(value) / maxValue * 100);
        const muted = label === 'Лиды' && !leadEnabled ? ' muted' : '';
        return `<div class="funnel-row${muted}"><div class="funnel-name">${label}</div><div class="funnel-track"><span style="width:${width.toFixed(2)}%"></span></div><div class="funnel-value"><strong>${label !== 'Лиды' || leadEnabled ? number(value) : 'нет данных'}</strong><small>${rate}</small></div></div>`;
      }).join('');
    }

    function renderLeadQuality(total, from, to) {
      const holder = document.getElementById('lead-quality');
      if (!holder) return;
      if (!leadEnabled) {
        holder.innerHTML = '<div class="quality-empty">нет данных</div>';
        return;
      }
      const quality = toNum(total.quality_leads);
      const bad = toNum(total.bad_leads);
      const base = quality + bad;
      if (!base) {
        holder.innerHTML = '<div class="quality-empty">Лидов за период нет</div>';
        return;
      }
      const rows = [
        ['Лиды', quality, quality / base * 100, 'quality-good'],
        ['Обращения', bad, bad / base * 100, 'quality-bad'],
      ];
      const summary = rows.map(([label, value, rate, cls]) => `<div class="quality-row ${cls}"><span>${label}</span><strong>${number(value)}</strong><small>${pct(rate, 1)}</small></div>`).join('');
      const stages = (reportData.leadStages || []).map((stage) => {
        const count = rowsInPeriod(stage.daily || [], from, to).reduce((sum, row) => sum + toNum(row.count), 0);
        return { ...stage, count };
      }).filter((stage) => stage.count > 0).sort((a, b) => toNum(b.count) - toNum(a.count));
      const stageHtml = stages.length
        ? `<div class="quality-stages"><div class="quality-stage-title">Стадии</div>${stages.slice(0, 6).map((stage) => {
            const cls = stage.quality === 'bad' ? 'quality-bad' : 'quality-good';
            const label = stage.quality === 'bad' ? 'обращение' : 'лид';
            return `<div class="quality-stage ${cls}"><span>${escapeHtml(stage.name || stage.id)}</span><strong>${number(stage.count)}</strong><small>${label}</small></div>`;
          }).join('')}</div>`
        : '';
      holder.innerHTML = summary + stageHtml;
    }

    function renderForecast(rows) {
      const result = forecast(rows);
      setText('forecast-confidence', `Надежность прогноза: ${result.confidence}`);
      const leadCard = leadEnabled ? number(result.projections[30].leads) : 'нет данных';
      document.getElementById('forecast-grid').innerHTML = `<article class="kpi-card accent-coral"><span class="kpi-label">7 дней</span><strong class="kpi-value">${number(result.projections[7].clicks)} кликов</strong><small>${money(result.projections[7].spend)}</small></article><article class="kpi-card accent-cyan"><span class="kpi-label">30 дней</span><strong class="kpi-value">${number(result.projections[30].clicks)} кликов</strong><small>${money(result.projections[30].spend)}</small></article><article class="kpi-card accent-lime"><span class="kpi-label">Лиды 30 дней</span><strong class="kpi-value">${leadCard}</strong><small></small></article>`;
    }

    function updatePeriodDisplay(from, to, rows) {
      const first = from || rows[0]?.date || reportData.dataStart;
      const last = to || rows[rows.length - 1]?.date || reportData.dataEnd;
      setText('overview-period-start', ruDate(first) || 'нет данных');
      setText('overview-period-end', ruDate(last));
    }

    function applySearch() {
      const query = (searchInput?.value || '').trim().toLowerCase();
      for (const row of document.querySelectorAll('.data-table tbody tr')) {
        row.classList.toggle('is-hidden', Boolean(query) && !row.textContent.toLowerCase().includes(query));
      }
    }

    function applyDateFilter() {
      let from = dateFromInput?.value || reportData.dataStart;
      let to = dateToInput?.value || reportData.dataEnd;
      if (from && to && from > to) [from, to] = [to, from];
      const campaigns = reportData.campaigns.map((campaign) => {
        const daily = rowsInPeriod(campaign.daily, from, to).map(complete);
        return { ...campaign, filteredDaily: daily, total: sumRows(daily) };
      }).sort((a, b) => toNum(b.total.clicks) - toNum(a.total.clicks));
      const daily = mergeDaily(reportData.campaigns, from, to);
      const total = sumRows(daily);

      updatePeriodDisplay(from, to, daily);
      setText('overview-active-days', number(activeRows(daily).length));
      setText('overview-impressions', number(total.impressions));
      setText('overview-clicks', number(total.clicks));
      setText('overview-leads', leadEnabled ? number(total.quality_leads) : 'нет данных');
      setText('overview-appeals', leadEnabled ? number(appealCount(total)) : 'нет данных');
      setText('overview-ctr', pct(total.ctr, 2));
      setText('overview-click-to-lead', leadEnabled ? pct(total.clickToLead, 2) : 'нет данных');
      setText('kpi-impressions', number(total.impressions));
      setText('kpi-clicks', number(total.clicks));
      setText('kpi-ctr', pct(total.ctr, 2));
      setText('kpi-leads', leadEnabled ? number(total.quality_leads) : 'нет данных');
      setText('kpi-appeals', leadEnabled ? number(appealCount(total)) : 'нет данных');
      setText('kpi-click-to-lead', leadEnabled ? pct(total.clickToLead, 2) : 'нет данных');
      setText('kpi-spend', money(total.spend));
      setText('kpi-spend-note', `CPC ${money(total.cpc)}`);

      document.getElementById('daily-chart').innerHTML = svgDaily(daily);
      renderCampaignTable(campaigns);
      renderFunnel(total);
      renderLeadQuality(total, from, to);
      renderForecast(daily);
      document.getElementById('groups-chart').innerHTML = svgBar(entityTotals(reportData.campaigns, 'groups', from, to), 'bar-blue', 7);
      document.getElementById('creatives-chart').innerHTML = svgBar(entityTotals(reportData.campaigns, 'creatives', from, to), 'bar-green', 7);
      applySearch();
    }

    searchInput?.addEventListener('input', applySearch);
    dateFromInput?.addEventListener('change', applyDateFilter);
    dateToInput?.addEventListener('change', applyDateFilter);
    document.getElementById('reset-period')?.addEventListener('click', () => {
      if (dateFromInput) dateFromInput.value = reportData.dataStart || '';
      if (dateToInput) dateToInput.value = reportData.dataEnd || '';
      applyDateFilter();
    });

    function syncNav() {
      const checkpoint = window.scrollY + 130;
      let active = sections[0];
      for (const section of sections) {
        if (section.offsetTop <= checkpoint) active = section;
      }
      for (const link of navLinks) {
        link.classList.toggle('is-active', active && `#${active.id}` === link.getAttribute('href'));
      }
    }

    applyDateFilter();
    syncNav();
    window.addEventListener('scroll', syncNav, { passive: true });
    window.addEventListener('scroll', hideChartTooltip, { passive: true });
    """

    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(REPORT_BRAND_NAME)}</title>
  <style>
    :root {{
      color-scheme: dark;
      --bg: #0d1117;
      --surface: #161b22;
      --surface-raised: #1c2128;
      --surface-soft: #21262d;
      --ink: #e6edf3;
      --ink-soft: #c9d1d9;
      --muted: #8b949e;
      --paper: var(--surface);
      --paper-warm: var(--bg);
      --mist: #1f6feb;
      --mist-strong: #30363d;
      --blue: #79c0ff;
      --blue-strong: #58a6ff;
      --green: #3fb950;
      --amber: #d29922;
      --red: #f85149;
      --violet: #a371f7;
      --lime: #7ee787;
      --line: rgba(240, 246, 252, 0.10);
      --line-strong: rgba(240, 246, 252, 0.18);
      --shadow: 0 24px 70px rgba(0, 0, 0, 0.38);
      --font-main: "Segoe UI", Arial, sans-serif;
      --max: 1320px;
      --radius: 6px;
    }}
    *, *::before, *::after {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      font-family: var(--font-main);
      color: var(--ink);
      background: var(--bg);
      overflow-x: hidden;
    }}
    ::selection {{ background: rgba(88, 166, 255, 0.28); }}
    a {{ color: inherit; text-decoration: none; }}
    .section-anchor {{ scroll-margin-top: 94px; }}
    .site-header {{
      position: fixed;
      z-index: 80;
      inset: 0 0 auto;
      display: grid;
      grid-template-columns: auto 1fr auto;
      align-items: center;
      gap: 28px;
      padding: 18px clamp(20px, 4vw, 60px);
      border-bottom: 1px solid var(--line);
      background: rgba(13, 17, 23, 0.84);
      backdrop-filter: blur(18px) saturate(120%);
    }}
    .brand {{ display: inline-flex; align-items: baseline; min-width: max-content; }}
    .brand__main {{ font-size: clamp(24px, 2.1vw, 34px); line-height: 1; font-weight: 400; }}
    .site-nav {{ display: flex; justify-content: center; align-items: center; gap: 6px; }}
    .site-nav a {{
      display: inline-flex;
      align-items: center;
      min-height: 40px;
      padding: 0 16px;
      border: 1px solid transparent;
      border-radius: 999px;
      color: var(--muted);
      font-size: 14px;
    }}
    .site-nav a:hover,
    .site-nav a.is-active {{ color: var(--ink); border-color: var(--line-strong); background: rgba(88, 166, 255, 0.12); }}
    .header-pill {{
      display: inline-flex;
      align-items: center;
      min-height: 38px;
      padding: 0 14px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface-soft);
      color: var(--ink-soft);
      font-size: 13px;
      white-space: nowrap;
    }}
    .overview-strip {{
      width: min(var(--max), calc(100% - 40px));
      margin: 104px auto 0;
      padding: 24px;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      background: linear-gradient(112deg, #161b22 0%, #1c2128 60%, #161b22 100%);
    }}
    .eyebrow {{ margin: 0 0 10px; color: var(--muted); font-size: 12px; text-transform: uppercase; }}
    h1, h2, h3, p {{ margin-top: 0; letter-spacing: 0; }}
    .overview-strip h1 {{ margin: 0; font-size: clamp(28px, 3.6vw, 40px); line-height: 1.02; font-weight: 500; }}
    .overview-summary-row {{
      display: grid;
      grid-auto-flow: column;
      grid-auto-columns: minmax(130px, 1fr);
      gap: 12px;
      margin-top: 22px;
      overflow-x: auto;
      padding-bottom: 4px;
    }}
    .overview-stat, .kpi-card, .panel {{
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface);
      box-shadow: var(--shadow);
    }}
    .overview-stat {{ min-height: 76px; padding: 12px 14px; box-shadow: none; background: var(--surface-raised); }}
    .overview-stat span, .kpi-label {{ display: block; color: var(--muted); font-size: 12px; }}
    .overview-stat strong {{ display: block; margin-top: 5px; color: var(--ink); font-size: clamp(16px, 1.45vw, 22px); line-height: 1.1; overflow-wrap: anywhere; }}
    .overview-stat--period {{ min-width: 210px; }}
    .period-value {{
      display: flex !important;
      flex-wrap: wrap;
      align-items: baseline;
      gap: 4px 8px;
      color: var(--ink);
      font-size: 18px !important;
      line-height: 1.18 !important;
    }}
    .period-value span {{ color: inherit; font-size: inherit; white-space: nowrap; }}
    .period-separator {{ color: var(--muted) !important; }}
    .page-shell {{ width: min(var(--max), calc(100% - 40px)); margin: 20px auto 32px; }}
    .filters-band, .kpi-grid, .chart-grid, .table-grid {{ margin-top: 20px; }}
    .filters-band {{
      display: grid;
      grid-template-columns: minmax(140px, 170px) minmax(140px, 170px) minmax(260px, 1fr) auto;
      gap: 12px;
      align-items: end;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface);
      box-shadow: var(--shadow);
    }}
    .field {{ display: grid; gap: 6px; min-width: 0; }}
    .field span {{ color: var(--muted); font-size: 12px; }}
    .field input {{
      min-height: 42px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--paper);
      color: var(--ink);
      color-scheme: dark;
      padding: 0 12px;
      min-width: 0;
    }}
    .field input::placeholder {{ color: #6e7681; }}
    .field input:focus {{
      outline: none;
      border-color: rgba(88, 166, 255, 0.56);
      box-shadow: 0 0 0 4px rgba(88, 166, 255, 0.16);
    }}
    .ghost-button {{
      min-height: 42px;
      padding: 0 16px;
      border: 1px solid var(--line-strong);
      border-radius: var(--radius);
      background: var(--surface-soft);
      color: var(--ink);
      cursor: pointer;
    }}
    .ghost-button:hover {{ border-color: rgba(88, 166, 255, 0.5); background: rgba(88, 166, 255, 0.14); }}
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 14px;
    }}
    .kpi-card {{
      min-height: 124px;
      padding: 16px;
      display: grid;
      gap: 12px;
      align-content: start;
      position: relative;
      overflow: hidden;
    }}
    .kpi-card::before {{
      content: "";
      position: absolute;
      inset: 0 auto auto 0;
      width: 100%;
      height: 4px;
      background: currentColor;
    }}
    .kpi-card small {{ color: var(--muted); min-height: 16px; }}
    .kpi-value {{ color: var(--ink); font-size: clamp(26px, 2.3vw, 34px); line-height: 1; font-weight: 700; }}
    .accent-cyan {{ color: var(--blue-strong); }}
    .accent-emerald {{ color: var(--green); }}
    .accent-amber {{ color: var(--amber); }}
    .accent-coral {{ color: var(--red); }}
    .accent-lime {{ color: var(--lime); }}
    .accent-violet {{ color: var(--violet); }}
    .chart-grid, .table-grid {{
      display: grid;
      grid-template-columns: repeat(12, minmax(0, 1fr));
      gap: 20px;
    }}
    .panel {{ grid-column: span 6; padding: 18px; min-width: 0; }}
    .panel-wide {{ grid-column: span 12; }}
    .panel-inner {{ box-shadow: none; margin-top: 16px; }}
    .panel-head {{ margin-bottom: 16px; }}
    .panel-head h2, .panel-head h3 {{ margin: 0; font-size: 19px; font-weight: 500; }}
    .panel-head h3 {{ font-size: 17px; }}
    .panel-head p {{ margin: 6px 0 0; color: var(--muted); font-size: 13px; }}
    .panel-head-outside {{ margin: 24px 0 12px; }}
    .chart-frame {{ position: relative; width: 100%; }}
    .chart-tooltip {{
      position: fixed;
      z-index: 120;
      display: grid;
      gap: 4px;
      max-width: min(280px, calc(100vw - 24px));
      padding: 10px 12px;
      border: 1px solid var(--line-strong);
      border-radius: var(--radius);
      background: #0d1117;
      color: var(--ink-soft);
      box-shadow: 0 16px 38px rgba(0, 0, 0, 0.42);
      font-size: 12px;
      line-height: 1.35;
      pointer-events: none;
    }}
    .chart-tooltip[hidden] {{ display: none; }}
    .chart-tooltip strong {{ color: var(--ink); font-size: 13px; }}
    .chart-tooltip span {{ color: var(--ink-soft); }}
    .forecast-grid, .mini-kpis {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }}
    .mini-kpis {{ grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); margin: 16px 0; }}
    .funnel-row {{
      display: grid;
      grid-template-columns: 84px minmax(140px, 1fr) 126px;
      gap: 12px;
      align-items: center;
      padding: 12px 0;
      border-bottom: 1px solid var(--line);
    }}
    .funnel-row:last-child {{ border-bottom: 0; }}
    .funnel-name {{ font-size: 14px; font-weight: 650; }}
    .funnel-track {{ height: 13px; border-radius: 999px; background: var(--surface-soft); overflow: hidden; }}
    .funnel-track span {{ display: block; height: 100%; border-radius: 999px; background: var(--blue-strong); }}
    .funnel-row:nth-child(2) .funnel-track span {{ background: var(--green); }}
    .funnel-row:nth-child(3) .funnel-track span {{ background: var(--amber); }}
    .funnel-row.muted .funnel-track span {{ background: #6e7681; }}
    .funnel-value strong {{ display: block; font-size: 15px; }}
    .funnel-value small {{ color: var(--muted); font-size: 12px; }}
    .quality-list {{ display: grid; gap: 10px; }}
    .quality-row {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto auto;
      gap: 12px;
      align-items: center;
      min-height: 48px;
      padding: 12px 0;
      border-bottom: 1px solid var(--line);
    }}
    .quality-row:last-child {{ border-bottom: 0; }}
    .quality-row span {{ color: var(--ink-soft); font-weight: 650; }}
    .quality-row strong {{ color: var(--ink); font-size: 18px; }}
    .quality-row small {{ color: var(--muted); min-width: 54px; text-align: right; }}
    .quality-good strong {{ color: var(--lime); }}
    .quality-bad strong {{ color: var(--red); }}
    .quality-empty {{ color: var(--muted); padding: 8px 0; }}
    .quality-stages {{
      display: grid;
      gap: 8px;
      margin-top: 8px;
      padding-top: 12px;
      border-top: 1px solid var(--line);
    }}
    .quality-stage-title {{ color: var(--muted); font-size: 12px; }}
    .quality-stage {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 4px 12px;
      align-items: baseline;
    }}
    .quality-stage span {{ color: var(--ink-soft); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .quality-stage strong {{ color: var(--ink); }}
    .quality-stage small {{ grid-column: 1 / -1; color: var(--muted); font-size: 12px; }}
    .table-wrap {{ overflow: auto; }}
    .data-table {{ width: 100%; border-collapse: collapse; min-width: 760px; }}
    .data-table th, .data-table td {{ padding: 11px 12px; border-bottom: 1px solid var(--line); vertical-align: middle; white-space: nowrap; }}
    .data-table th {{ position: sticky; top: 0; z-index: 1; background: var(--surface-raised); color: var(--muted); font-size: 12px; text-align: left; }}
    .data-table td {{ font-size: 14px; }}
    .data-table tbody tr:hover {{ background: rgba(88, 166, 255, 0.08); }}
    .is-hidden {{ display: none !important; }}
    .num {{ text-align: right; }}
    .empty-row td {{ padding: 28px 12px; color: var(--muted); text-align: center; }}
    svg {{ display: block; width: 100%; height: auto; }}
    .grid-line {{ stroke: rgba(240, 246, 252, 0.10); stroke-width: 1; }}
    .axis-label, .axis-title, .bar-label, .bar-value, .svg-muted, .chart-legend text {{ fill: var(--muted); font-size: 12px; }}
    .axis-title {{ fill: var(--ink-soft); font-weight: 650; }}
    .bar-label {{ fill: var(--ink-soft); }}
    .bar-value {{ fill: var(--ink); font-weight: 650; }}
    .bar-clicks, .bar-blue {{ fill: var(--blue-strong); }}
    .bar-green {{ fill: var(--green); }}
    .line-ctr {{ stroke: var(--violet); stroke-width: 3; stroke-linecap: round; stroke-linejoin: round; }}
    .dot-ctr {{ fill: var(--violet); }}
    .chart-hover {{ fill: transparent; pointer-events: all; cursor: crosshair; }}
    .campaign {{
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: var(--surface);
      box-shadow: var(--shadow);
      margin-bottom: 14px;
      overflow: hidden;
    }}
    .campaign summary {{
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      gap: 16px;
      padding: 16px 18px;
      list-style: none;
    }}
    .campaign summary::-webkit-details-marker {{ display: none; }}
    .campaign-title {{ font-weight: 650; }}
    .campaign-title em {{ color: var(--muted); font-style: normal; font-weight: 400; margin-left: 6px; }}
    .campaign summary > span:last-child {{ color: var(--muted); text-align: right; }}
    .campaign-body {{ padding: 0 18px 18px; border-top: 1px solid var(--line); }}
    footer {{ width: min(var(--max), calc(100% - 40px)); margin: 24px auto 34px; color: var(--muted); font-size: 12px; }}
    @media (max-width: 1240px) {{
      .kpi-grid {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
      .panel, .panel-wide {{ grid-column: span 12; }}
      .mini-kpis {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    }}
    @media (max-width: 860px) {{
      .site-header {{ grid-template-columns: 1fr; justify-items: start; gap: 12px; position: static; }}
      .site-nav {{ flex-wrap: wrap; justify-content: flex-start; }}
      .filters-band {{ grid-template-columns: 1fr; }}
      .overview-strip {{ margin-top: 20px; }}
      .overview-summary-row {{ grid-auto-columns: minmax(150px, 1fr); }}
      .page-shell, .overview-strip, footer {{ width: min(100% - 20px, var(--max)); }}
      .funnel-row {{ grid-template-columns: 1fr; gap: 6px; }}
    }}
    @media (max-width: 720px) {{
      .overview-strip {{ padding: 18px; }}
      .overview-summary-row {{ grid-auto-flow: row; grid-auto-columns: unset; overflow-x: visible; }}
      .kpi-grid, .forecast-grid, .mini-kpis {{ grid-template-columns: 1fr; }}
      .kpi-value {{ font-size: 30px; }}
      .campaign summary {{ display: block; }}
      .campaign summary > span:last-child {{ display: block; margin-top: 6px; text-align: left; }}
    }}
    @media print {{
      .site-header {{ position: static; }}
      .panel, .campaign, .kpi-card, .overview-strip {{ box-shadow: none; break-inside: avoid; }}
    }}
  </style>
</head>
<body>
  <header class="site-header">
      <a class="brand" href="#top"><span class="brand__main">{esc(REPORT_BRAND_NAME)}</span></a>
    <nav class="site-nav" aria-label="Разделы">
      <a href="#overview">Обзор</a>
      <a href="#campaigns">Кампании</a>
      <a href="#details">Детализация</a>
    </nav>
    <div class="header-pill">{esc("Лиды подключены" if lead_enabled else "Лиды: нет данных")}</div>
  </header>

  <main id="top">
    <section class="overview-strip section-anchor" id="overview">
      <p class="eyebrow">{esc(REPORT_EYEBROW)}</p>
      <h1>{esc(REPORT_HEADING)}</h1>
      <div class="overview-summary-row" aria-label="Краткая сводка">
        {overview_stat("Кампаний", number(len(campaigns)), "campaigns")}
        {overview_period_stat(report["data_start"], report["data_end"])}
        {overview_stat("Активных дней", number(report["active_days"]), "active-days")}
        {overview_stat("Показы", number(total["impressions"]), "impressions")}
        {overview_stat("Клики", number(total["clicks"]), "clicks")}
        {overview_stat("Лиды", number(qualified_leads(total)) if lead_enabled else "нет данных", "leads")}
        {overview_stat("Обращения", number(appeal_count(total)) if lead_enabled else "нет данных", "appeals")}
        {overview_stat("Показ → клик", pct(total["ctr"], 2), "ctr")}
        {overview_stat("Клик → лид", pct(total["click_to_lead"], 2) if lead_enabled else "нет данных", "click-to-lead")}
      </div>
    </section>

    <div class="page-shell">
      {main_content}
    </div>
  </main>

  <footer>Сгенерировано {datetime.now().strftime("%Y-%m-%d %H:%M")}</footer>
  <div id="chart-tooltip" class="chart-tooltip" hidden></div>
  <script id="report-data" type="application/json">{report_data_json}</script>
  <script>
{client_script}
  </script>
</body>
</html>
"""


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    report = build_report()
    html_report = render_html(report)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    INDEX_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_text(html_report, encoding="utf-8")
    INDEX_FILE.write_text(html_report, encoding="utf-8")
    print(f"Готово: {OUTPUT_FILE}")
    print(f"Готово: {INDEX_FILE}")
    print(f"Кампаний: {len(report['campaigns'])}. Лиды: {report['lead_data'].status}")


if __name__ == "__main__":
    main()
