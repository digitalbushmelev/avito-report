# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from urllib import parse, request
from zoneinfo import ZoneInfo


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


AVITO_ENV_FILE = path_from_env("REPORT_AVITO_ENV_FILE", "avito.env")
AVITO_CAMPAIGN_REGISTRY_FILE = path_from_env("REPORT_AVITO_CAMPAIGN_REGISTRY_FILE", "avito_campaigns.json")
AVITO_DATA_DIR = path_from_env("REPORT_AVITO_DATA_DIR", "data/avito")
AVITO_RAW_DIR = path_from_env("REPORT_AVITO_RAW_DIR", str(AVITO_DATA_DIR / "raw"))
REPORT_SCRIPT = path_from_env("REPORT_SCRIPT", "generate_avito_report.py")
REPORT_TIMEZONE = "Asia/Yekaterinburg"
DEFAULT_AVITO_STATS_PATH = "/ads/v1/account/{account_id}/campaigns/{campaign_id}/stats"
DEFAULT_AVITO_CAMPAIGN_LIST_PATH = "/ads/v1/account/{account_id}/campaigns"
DEFAULT_AVITO_CREATIVE_LIST_PATH = "/ads/v1/account/{account_id}/creatives"


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


def split_list(value: object) -> list[str]:
    return [item.strip() for item in text(value).split(",") if item.strip()]


def slugify(value: object, fallback: str = "campaign") -> str:
    translit = {
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "е": "e",
        "ё": "e",
        "ж": "zh",
        "з": "z",
        "и": "i",
        "й": "y",
        "к": "k",
        "л": "l",
        "м": "m",
        "н": "n",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "у": "u",
        "ф": "f",
        "х": "h",
        "ц": "c",
        "ч": "ch",
        "ш": "sh",
        "щ": "sch",
        "ъ": "",
        "ы": "y",
        "ь": "",
        "э": "e",
        "ю": "yu",
        "я": "ya",
    }
    raw = text(value).lower()
    latin = "".join(translit.get(char, char) for char in raw)
    slug = re.sub(r"[^a-z0-9]+", "-", latin).strip("-")
    return slug or fallback


def read_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if path.exists():
        for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    for key, value in os.environ.items():
        if key.startswith(("AVITO_", "REPORT_")):
            values[key] = value
    return values


def now_local() -> datetime:
    return datetime.now(ZoneInfo(REPORT_TIMEZONE))


def cache_is_fresh(path: Path) -> bool:
    if not path.exists():
        return False
    updated_at = datetime.fromtimestamp(path.stat().st_mtime, tz=ZoneInfo(REPORT_TIMEZONE))
    return updated_at.date() == now_local().date()


def date_range(env: dict[str, str]) -> tuple[str, str]:
    today = now_local().date()
    date_to = text(env.get("AVITO_DATE_TO")) or (today - timedelta(days=1)).isoformat()
    if text(env.get("AVITO_DATE_FROM")):
        return text(env["AVITO_DATE_FROM"]), date_to
    lookback = int(text(env.get("AVITO_LOOKBACK_DAYS")) or "90")
    date_from = (datetime.fromisoformat(date_to).date() - timedelta(days=max(1, lookback) - 1)).isoformat()
    return date_from, date_to


def avito_token(env: dict[str, str]) -> str:
    token_url = text(env.get("AVITO_TOKEN_URL")) or "https://api.avito.ru/token"
    client_id = text(env.get("AVITO_CLIENT_ID") or env.get("AVITO_CLIENT_KEY"))
    client_secret = text(env.get("AVITO_CLIENT_SECRET"))
    if not client_id or not client_secret:
        raise RuntimeError("Заполните AVITO_CLIENT_ID и AVITO_CLIENT_SECRET в avito.env")

    payload = parse.urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
        }
    ).encode("utf-8")
    req = request.Request(token_url, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with request.urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode("utf-8"))
    token = text(data.get("access_token"))
    if not token:
        raise RuntimeError("Avito не вернул access_token")
    return token


def format_template(
    value: str,
    env: dict[str, str],
    date_from: str = "",
    date_to: str = "",
    campaign: dict[str, str] | None = None,
) -> str:
    campaign = campaign or {}
    return value.format(
        account_id=text(env.get("AVITO_ACCOUNT_ID")),
        campaign_id=text(campaign.get("id") or env.get("AVITO_CAMPAIGN_ID")),
        campaign_name=text(campaign.get("name") or env.get("AVITO_CAMPAIGN_NAME") or "Дружеский"),
        date_from=date_from,
        date_to=date_to,
        group_by=text(env.get("AVITO_GROUP_BY") or "day"),
    )


def avito_json_request(
    env: dict[str, str],
    token: str,
    method: str,
    path_template: str,
    date_from: str = "",
    date_to: str = "",
    campaign: dict[str, str] | None = None,
    body_template: str = "",
) -> dict[str, object]:
    base_url = (text(env.get("AVITO_BASE_URL")) or "https://api.avito.ru").rstrip("/")
    method = method.upper()
    url = base_url + format_template(path_template, env, date_from, date_to, campaign)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    body = None
    if body_template:
        body_text = format_template(body_template, env, date_from, date_to, campaign)
        body = json.dumps(json.loads(body_text), ensure_ascii=False).encode("utf-8")
    elif method == "POST" and date_from and date_to:
        body = json.dumps({"dateFrom": date_from, "dateTo": date_to}, ensure_ascii=False).encode("utf-8")
    elif method == "GET":
        query = parse.urlencode({"dateFrom": date_from, "dateTo": date_to, "groupBy": text(env.get("AVITO_GROUP_BY") or "day")})
        url += ("&" if "?" in url else "?") + query

    req = request.Request(url, data=body, headers=headers, method=method)
    with request.urlopen(req, timeout=45) as response:
        data = json.loads(response.read().decode("utf-8"))
        if isinstance(data, dict):
            balance = response.headers.get("Api-Point-Balance")
            if balance:
                data["_api_point_balance"] = balance
        return data


def avito_stats_request(env: dict[str, str], token: str, date_from: str, date_to: str, campaign: dict[str, str]) -> dict[str, object]:
    return avito_json_request(
        env=env,
        token=token,
        method=text(env.get("AVITO_STATS_METHOD")) or "POST",
        path_template=text(env.get("AVITO_STATS_PATH_TEMPLATE")) or DEFAULT_AVITO_STATS_PATH,
        date_from=date_from,
        date_to=date_to,
        campaign=campaign,
        body_template=text(env.get("AVITO_STATS_BODY_JSON")),
    )


def first_number(row: dict[str, object], names: tuple[str, ...]) -> float:
    for name in names:
        if name in row and row[name] not in ("", None):
            try:
                return float(str(row[name]).replace(",", "."))
            except ValueError:
                return 0.0
    return 0.0


def metric_row(row: dict[str, object]) -> dict[str, float]:
    return {
        "impressions": first_number(row, ("impressions", "views", "shows", "show_count")),
        "clicks": first_number(row, ("clicks", "click_count")),
        "spend": first_number(row, ("spend", "cost", "expense", "amount")),
        "bonus_spend": first_number(row, ("bonus_spend", "bonusSpend", "spendBonus", "bonus")),
    }


def row_date(row: dict[str, object]) -> str:
    return text(row.get("date") or row.get("day") or row.get("period") or row.get("timestamp"))[:10]


def normalize_data_points(items: object) -> list[dict[str, object]]:
    daily = []
    for row in items if isinstance(items, list) else []:
        if not isinstance(row, dict):
            continue
        day = row_date(row)
        if not day:
            continue
        daily.append({"date": day, **metric_row(row)})
    return sorted(daily, key=lambda item: text(item.get("date")))


def normalize_avito_entities(items: object) -> list[dict[str, object]]:
    entities = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        label = text(item.get("id") or item.get("name") or item.get("label"))
        daily = normalize_data_points(item.get("data"))
        total_source = item.get("totalData") if isinstance(item.get("totalData"), dict) else item
        entities.append(
            {
                "id": text(item.get("id") or label),
                "label": text(item.get("name") or label),
                "group_id": text(item.get("groupId") or item.get("groupID")),
                "daily": daily,
                **metric_row(total_source if isinstance(total_source, dict) else {}),
            }
        )
    return entities


def find_rows(value: object) -> list[dict[str, object]]:
    if isinstance(value, list):
        direct = [item for item in value if isinstance(item, dict) and any(key in item for key in ("date", "day", "period"))]
        if direct:
            return direct
        rows: list[dict[str, object]] = []
        for item in value:
            rows.extend(find_rows(item))
        return rows
    if isinstance(value, dict):
        for key in ("daily", "statistics", "stats", "rows", "items", "data", "result"):
            if key in value:
                rows = find_rows(value[key])
                if rows:
                    return rows
        rows = []
        for item in value.values():
            rows.extend(find_rows(item))
        return rows
    return []


def campaign_id_from_item(item: dict[str, object]) -> str:
    return text(item.get("id") or item.get("campaignId") or item.get("campaignID") or item.get("campaign_id"))


def campaign_name_from_item(item: dict[str, object]) -> str:
    return text(item.get("name") or item.get("title") or item.get("campaignName") or item.get("campaign_name"))


def campaign_slug(campaign: dict[str, str]) -> str:
    if text(campaign.get("slug")):
        return slugify(campaign.get("slug"), text(campaign.get("id")) or "campaign")
    if text(campaign.get("id")) == text(campaign.get("legacy_id")) and text(campaign.get("legacy_slug")):
        return text(campaign.get("legacy_slug"))
    return slugify(campaign.get("name") or campaign.get("id"), text(campaign.get("id")) or "campaign")


def campaign_data_file(campaign: dict[str, str]) -> Path:
    return AVITO_DATA_DIR / f"{campaign_slug(campaign)}.json"


def campaigns_from_items(items: object) -> list[dict[str, str]]:
    campaigns: list[dict[str, str]] = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        campaign_id = text(item.get("id") or item.get("campaign_id") or item.get("campaignId"))
        if not campaign_id:
            continue
        campaign = {
            "id": campaign_id,
            "name": text(item.get("name") or item.get("title") or campaign_id),
            "slug": text(item.get("slug")),
            "lead_date_from": text(item.get("lead_date_from") or item.get("date_from")),
            "lead_date_to": text(item.get("lead_date_to") or item.get("date_to")),
        }
        lead_keys = item.get("lead_keys") or item.get("utm_campaigns") or item.get("utm_campaign")
        keys = split_list(lead_keys) if not isinstance(lead_keys, list) else [text(value) for value in lead_keys if text(value)]
        if keys:
            campaign["lead_keys"] = ",".join(dict.fromkeys(keys))
        campaigns.append(campaign)
    return campaigns


def campaign_registry_path(env: dict[str, str]) -> Path:
    value = text(env.get("AVITO_CAMPAIGN_REGISTRY_FILE"))
    if not value:
        return AVITO_CAMPAIGN_REGISTRY_FILE
    path = Path(value)
    return path if path.is_absolute() else ROOT / path


def registry_campaigns(env: dict[str, str]) -> list[dict[str, str]]:
    path = campaign_registry_path(env)
    if not path.exists():
        return []
    try:
        parsed = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"Avito campaign registry пропущен: {path.name}: {exc}")
        return []
    items = parsed.get("campaigns") if isinstance(parsed, dict) else parsed
    campaigns = campaigns_from_items(items)
    if campaigns:
        print(f"Avito campaign registry: кампаний {len(campaigns)}")
    return campaigns


def configured_campaigns(env: dict[str, str]) -> list[dict[str, str]]:
    campaigns: list[dict[str, str]] = registry_campaigns(env)
    raw_json = text(env.get("AVITO_CAMPAIGNS_JSON"))
    if raw_json:
        try:
            parsed = json.loads(raw_json)
        except json.JSONDecodeError:
            parsed = []
        campaigns.extend(campaigns_from_items(parsed))

    ids = split_list(env.get("AVITO_CAMPAIGN_IDS"))
    if ids:
        names = split_list(env.get("AVITO_CAMPAIGN_NAMES"))
        slugs = split_list(env.get("AVITO_CAMPAIGN_SLUGS"))
        for idx, campaign_id in enumerate(ids):
            campaigns.append(
                {
                    "id": campaign_id,
                    "name": names[idx] if idx < len(names) else campaign_id,
                    "slug": slugs[idx] if idx < len(slugs) else "",
                }
            )

    legacy_id = text(env.get("AVITO_CAMPAIGN_ID"))
    if legacy_id:
        campaigns.append(
            {
                "id": legacy_id,
                "name": text(env.get("AVITO_CAMPAIGN_NAME") or "Дружеский"),
                "slug": text(env.get("AVITO_CAMPAIGN_SLUG") or "druzheskiy"),
                "legacy_id": legacy_id,
                "legacy_slug": "druzheskiy",
            }
        )

    result: dict[str, dict[str, str]] = {}
    for campaign in campaigns:
        result[campaign["id"]] = {**result.get(campaign["id"], {}), **campaign}
    return list(result.values())


def extract_campaigns(value: object) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []
    if isinstance(value, list):
        for item in value:
            candidates.extend(extract_campaigns(item))
        return candidates
    if not isinstance(value, dict):
        return candidates

    campaign_id = campaign_id_from_item(value)
    if campaign_id:
        candidates.append({"id": campaign_id, "name": campaign_name_from_item(value) or campaign_id})

    for key in ("campaigns", "items", "result", "data"):
        if key in value:
            candidates.extend(extract_campaigns(value[key]))
    return candidates


def auto_discover_enabled(env: dict[str, str]) -> bool:
    return text(env.get("AVITO_AUTO_DISCOVER_CAMPAIGNS") or "1").lower() not in {"0", "false", "no", "off"}


def registry_write_enabled(env: dict[str, str]) -> bool:
    return text(env.get("AVITO_WRITE_DISCOVERED_CAMPAIGNS") or "1").lower() not in {"0", "false", "no", "off"}


def creative_discovery_enabled(env: dict[str, str]) -> bool:
    return text(env.get("AVITO_DISCOVER_CAMPAIGNS_FROM_CREATIVES") or "1").lower() not in {"0", "false", "no", "off"}


def save_campaign_registry(env: dict[str, str], campaigns: list[dict[str, str]]) -> None:
    if not registry_write_enabled(env):
        return
    path = campaign_registry_path(env)
    rows = []
    existing = {campaign["id"]: campaign for campaign in registry_campaigns(env)}
    seen: set[str] = set()
    for campaign in campaigns:
        campaign_id = text(campaign.get("id"))
        if not campaign_id or campaign_id in seen:
            continue
        seen.add(campaign_id)
        current = {**existing.get(campaign_id, {}), **campaign}
        row = {
            "id": campaign_id,
            "name": text(current.get("name") or campaign_id),
            "slug": campaign_slug(current),
        }
        if text(current.get("lead_date_from")):
            row["lead_date_from"] = text(current.get("lead_date_from"))
        if text(current.get("lead_date_to")):
            row["lead_date_to"] = text(current.get("lead_date_to"))
        lead_keys = split_list(current.get("lead_keys"))
        if lead_keys:
            row["lead_keys"] = list(dict.fromkeys(lead_keys))
        rows.append(row)
    if not rows:
        return
    payload = {"campaigns": rows}
    new_text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    old_text = path.read_text(encoding="utf-8") if path.exists() else ""
    if new_text != old_text:
        path.write_text(new_text, encoding="utf-8")
        print(f"Avito campaign registry обновлен: {path.name}")


def first_active_day(payload: dict[str, object]) -> str:
    rows = payload.get("daily") if isinstance(payload.get("daily"), list) else []
    for row in rows:
        if not isinstance(row, dict):
            continue
        if first_number(row, ("impressions",)) > 0 or first_number(row, ("clicks",)) > 0 or first_number(row, ("spend",)) > 0:
            return text(row.get("date"))
    for row in rows:
        if isinstance(row, dict) and text(row.get("date")):
            return text(row.get("date"))
    return ""


def update_registry_lead_windows(env: dict[str, str], first_days: dict[str, str]) -> None:
    path = campaign_registry_path(env)
    if not path.exists():
        return
    try:
        parsed = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return
    items = parsed.get("campaigns") if isinstance(parsed, dict) else parsed
    if not isinstance(items, list):
        return

    changed = False
    for item in items:
        if not isinstance(item, dict):
            continue
        campaign_id = text(item.get("id"))
        if campaign_id in first_days and not text(item.get("lead_date_from") or item.get("date_from")):
            item["lead_date_from"] = first_days[campaign_id]
            changed = True

    dated_items = []
    for item in items:
        if not isinstance(item, dict):
            continue
        start = text(item.get("lead_date_from") or item.get("date_from"))
        if start:
            dated_items.append((start, item))
    dated_items.sort(key=lambda pair: pair[0])

    for index, (start, item) in enumerate(dated_items[:-1]):
        if text(item.get("lead_date_to") or item.get("date_to")):
            continue
        next_start = datetime.fromisoformat(dated_items[index + 1][0]).date()
        date_to = (next_start - timedelta(days=1)).isoformat()
        if date_to >= start:
            item["lead_date_to"] = date_to
            changed = True

    if changed:
        path.write_text(json.dumps({"campaigns": items}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"Avito campaign registry: даты лидов обновлены в {path.name}")


def discover_campaigns(env: dict[str, str], token: str) -> list[dict[str, str]]:
    if not auto_discover_enabled(env):
        return []
    path = text(env.get("AVITO_CAMPAIGN_LIST_PATH_TEMPLATE")) or DEFAULT_AVITO_CAMPAIGN_LIST_PATH
    method = text(env.get("AVITO_CAMPAIGN_LIST_METHOD") or "GET")
    try:
        raw = avito_json_request(env, token, method, path)
    except Exception as exc:
        print(f"Avito campaign list пропущен: {exc}")
        return []

    result: dict[str, dict[str, str]] = {}
    for campaign in extract_campaigns(raw):
        if campaign["id"]:
            result[campaign["id"]] = campaign
    if result:
        print(f"Avito campaign list: найдено кампаний {len(result)}")
    return list(result.values())


def discover_campaigns_from_creatives(env: dict[str, str], token: str) -> list[dict[str, str]]:
    if not auto_discover_enabled(env) or not creative_discovery_enabled(env):
        return []

    base_url = (text(env.get("AVITO_BASE_URL")) or "https://api.avito.ru").rstrip("/")
    path = text(env.get("AVITO_CREATIVE_LIST_PATH_TEMPLATE")) or DEFAULT_AVITO_CREATIVE_LIST_PATH
    method = text(env.get("AVITO_CREATIVE_LIST_METHOD") or "POST").upper()
    limit = int(text(env.get("AVITO_CREATIVE_LIST_LIMIT")) or "100")
    body_template = text(env.get("AVITO_CREATIVE_LIST_BODY_JSON"))
    page = 1
    seen_creatives = 0
    campaigns: dict[str, dict[str, object]] = {}

    while True:
        url = base_url + format_template(path, env)
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        body = None
        if method == "POST":
            if body_template:
                body_text = body_template.format(
                    account_id=text(env.get("AVITO_ACCOUNT_ID")),
                    page=page,
                    limit=limit,
                )
                body_payload = json.loads(body_text)
            else:
                body_payload = {"limit": limit, "page": page}
            body = json.dumps(body_payload, ensure_ascii=False).encode("utf-8")
        elif method == "GET":
            query = parse.urlencode({"limit": limit, "page": page})
            url += ("&" if "?" in url else "?") + query

        req = request.Request(url, data=body, headers=headers, method=method)
        try:
            with request.urlopen(req, timeout=45) as response:
                raw = json.loads(response.read().decode("utf-8"))
        except Exception as exc:
            print(f"Avito creative list пропущен: {exc}")
            return []

        creatives = raw.get("creatives") if isinstance(raw, dict) else []
        if not isinstance(creatives, list) or not creatives:
            break

        seen_creatives += len(creatives)
        for creative in creatives:
            if not isinstance(creative, dict):
                continue
            campaign_id = text(
                creative.get("campaignID")
                or creative.get("campaignId")
                or creative.get("campaign_id")
            )
            if not campaign_id:
                continue
            campaign = campaigns.setdefault(
                campaign_id,
                {"id": campaign_id, "name": campaign_id, "lead_keys": []},
            )
            link = text(creative.get("link"))
            if link:
                for value in parse.parse_qs(parse.urlparse(link).query).get("utm_campaign", []):
                    key = text(value)
                    if key and "{" not in key and "}" not in key and key not in campaign["lead_keys"]:
                        campaign["lead_keys"].append(text(value))

        total = int(raw.get("total") or 0) if isinstance(raw, dict) else 0
        if total and seen_creatives >= total:
            break
        page += 1

    result = []
    for campaign in campaigns.values():
        row = {
            "id": text(campaign.get("id")),
            "name": text(campaign.get("name") or campaign.get("id")),
        }
        lead_keys = [text(value) for value in campaign.get("lead_keys", []) if text(value)]
        if lead_keys:
            row["lead_keys"] = ",".join(lead_keys)
        result.append(row)

    if result:
        print(f"Avito creative list: найдено кампаний {len(result)}")
    return result


def resolve_campaigns(env: dict[str, str], token: str) -> list[dict[str, str]]:
    configured = configured_campaigns(env)
    discovered = discover_campaigns(env, token)
    creative_discovered = discover_campaigns_from_creatives(env, token)
    merged: dict[str, dict[str, str]] = {}
    for campaign in discovered + creative_discovered + configured:
        campaign_id = text(campaign.get("id"))
        if not campaign_id:
            continue
        merged[campaign_id] = {**merged.get(campaign_id, {}), **campaign}
    campaigns = list(merged.values()) or configured
    if discovered or creative_discovered:
        save_campaign_registry(env, campaigns)
    return campaigns


def normalize_avito_response(
    raw: dict[str, object],
    env: dict[str, str],
    date_from: str,
    date_to: str,
    campaign: dict[str, str],
) -> dict[str, object]:
    campaign_node = raw.get("campaign") if isinstance(raw.get("campaign"), dict) else {}
    daily = normalize_data_points(campaign_node.get("data") if isinstance(campaign_node, dict) else [])
    if not daily:
        daily = [{"date": row_date(row), **metric_row(row)} for row in find_rows(raw) if row_date(row)]

    if not daily:
        total_source = campaign_node.get("totalData") if isinstance(campaign_node.get("totalData"), dict) else raw
        daily = [{"date": date_to, **metric_row(total_source if isinstance(total_source, dict) else {})}]

    return {
        "source": "Avito API",
        "updated_at": now_local().isoformat(),
        "api_point_balance": text(raw.get("_api_point_balance")),
        "period": {"start": date_from, "end": date_to},
        "campaign": {
            "id": text(campaign_node.get("id") if isinstance(campaign_node, dict) else "") or text(campaign.get("id")),
            "name": text(campaign_node.get("name") if isinstance(campaign_node, dict) else "") or text(campaign.get("name") or campaign.get("id")),
            "lead_keys": split_list(campaign.get("lead_keys")),
            "lead_date_from": text(campaign.get("lead_date_from")),
            "lead_date_to": text(campaign.get("lead_date_to")),
        },
        "daily": daily,
        "groups": normalize_avito_entities(raw.get("groups")),
        "creatives": normalize_avito_entities(raw.get("creatives")),
        "raw_saved": True,
    }


def update_avito_cache(force: bool = False) -> bool:
    env = read_env_file(AVITO_ENV_FILE)
    if text(env.get("AVITO_ENABLE_API")) != "1":
        print("Avito API выключен: поставьте AVITO_ENABLE_API=1 в avito.env")
        return False

    date_from, date_to = date_range(env)
    token = avito_token(env)
    campaigns = resolve_campaigns(env, token)
    if not campaigns:
        raise RuntimeError("Avito: нет кампаний для обновления")

    AVITO_DATA_DIR.mkdir(parents=True, exist_ok=True)
    AVITO_RAW_DIR.mkdir(parents=True, exist_ok=True)
    updated = False
    active_starts: dict[str, str] = {}
    normalized_campaigns: list[dict[str, str]] = []
    for campaign in campaigns:
        path = campaign_data_file(campaign)
        if not force and cache_is_fresh(path):
            print(f"Avito API пропущен для {campaign.get('name') or campaign.get('id')}: данные уже обновлялись сегодня")
            continue

        raw = avito_stats_request(env, token, date_from, date_to, campaign)
        slug = campaign_slug(campaign)
        raw_path = AVITO_RAW_DIR / f"{slug}_{now_local().strftime('%Y%m%d_%H%M%S')}.json"
        raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
        normalized = normalize_avito_response(raw, env, date_from, date_to, campaign)
        normalized_campaign = normalized.get("campaign") if isinstance(normalized.get("campaign"), dict) else {}
        normalized_id = text(normalized_campaign.get("id") or campaign.get("id"))
        normalized_name = text(normalized_campaign.get("name") or campaign.get("name") or campaign.get("id"))
        normalized_slug = text(campaign.get("slug"))
        if not normalized_slug or normalized_slug == normalized_id:
            normalized_slug = campaign_slug({"id": normalized_id, "name": normalized_name})
        save_campaign = {
            **campaign,
            "id": normalized_id,
            "name": normalized_name,
            "slug": normalized_slug,
        }
        normalized_campaigns.append(save_campaign)
        save_path = campaign_data_file(save_campaign)
        if save_path != path and path.exists():
            path.unlink()
        save_path.write_text(
            json.dumps(normalized, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        active_day = first_active_day(normalized)
        if active_day:
            active_starts[text(campaign.get("id"))] = active_day
        print(f"Avito API обновлен: {save_path}")
        updated = True
    if active_starts:
        update_registry_lead_windows(env, active_starts)
    if normalized_campaigns:
        save_campaign_registry(env, normalized_campaigns)
    return updated


def rebuild_report() -> None:
    subprocess.run([sys.executable, str(REPORT_SCRIPT)], cwd=ROOT, check=True, env=os.environ.copy())


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true", help="принудительно запросить Avito API, даже если сегодня уже обновляли")
    parser.add_argument("--skip-avito", action="store_true", help="не дергать Avito API, только пересобрать отчет")
    args = parser.parse_args()

    if not args.skip_avito:
        update_avito_cache(force=args.force)
    rebuild_report()


if __name__ == "__main__":
    main()
