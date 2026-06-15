# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from urllib import parse, request
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parent
AVITO_ENV_FILE = ROOT / "avito.env"
AVITO_DATA_DIR = ROOT / "data" / "avito"
AVITO_RAW_DIR = AVITO_DATA_DIR / "raw"
AVITO_DATA_FILE = AVITO_DATA_DIR / "druzheskiy.json"
REPORT_SCRIPT = ROOT / "generate_avito_report.py"
REPORT_TIMEZONE = "Asia/Yekaterinburg"
DEFAULT_AVITO_STATS_PATH = "/ads/v1/account/{account_id}/campaigns/{campaign_id}/stats"


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


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


def format_template(value: str, env: dict[str, str], date_from: str, date_to: str) -> str:
    return value.format(
        account_id=text(env.get("AVITO_ACCOUNT_ID")),
        campaign_id=text(env.get("AVITO_CAMPAIGN_ID")),
        campaign_name=text(env.get("AVITO_CAMPAIGN_NAME") or "Дружеский"),
        date_from=date_from,
        date_to=date_to,
        group_by=text(env.get("AVITO_GROUP_BY") or "day"),
    )


def avito_stats_request(env: dict[str, str], token: str, date_from: str, date_to: str) -> dict[str, object]:
    base_url = (text(env.get("AVITO_BASE_URL")) or "https://api.avito.ru").rstrip("/")
    path_template = text(env.get("AVITO_STATS_PATH_TEMPLATE")) or DEFAULT_AVITO_STATS_PATH
    method = (text(env.get("AVITO_STATS_METHOD")) or "POST").upper()
    url = base_url + format_template(path_template, env, date_from, date_to)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    body_template = text(env.get("AVITO_STATS_BODY_JSON"))
    body = None
    if body_template:
        body_text = format_template(body_template, env, date_from, date_to)
        body = json.dumps(json.loads(body_text), ensure_ascii=False).encode("utf-8")
    elif method == "POST":
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


def normalize_avito_response(raw: dict[str, object], env: dict[str, str], date_from: str, date_to: str) -> dict[str, object]:
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
            "id": text(campaign_node.get("id") if isinstance(campaign_node, dict) else "") or text(env.get("AVITO_CAMPAIGN_ID") or "druzheskiy"),
            "name": text(campaign_node.get("name") if isinstance(campaign_node, dict) else "") or text(env.get("AVITO_CAMPAIGN_NAME") or "Дружеский"),
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
    if not force and cache_is_fresh(AVITO_DATA_FILE):
        print("Avito API пропущен: данные уже обновлялись сегодня")
        return False

    date_from, date_to = date_range(env)
    token = avito_token(env)
    raw = avito_stats_request(env, token, date_from, date_to)

    AVITO_DATA_DIR.mkdir(parents=True, exist_ok=True)
    AVITO_RAW_DIR.mkdir(parents=True, exist_ok=True)
    raw_path = AVITO_RAW_DIR / f"druzheskiy_{now_local().strftime('%Y%m%d_%H%M%S')}.json"
    raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    AVITO_DATA_FILE.write_text(json.dumps(normalize_avito_response(raw, env, date_from, date_to), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Avito API обновлен: {AVITO_DATA_FILE}")
    return True


def rebuild_report() -> None:
    subprocess.run([sys.executable, str(REPORT_SCRIPT)], cwd=ROOT, check=True)


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
