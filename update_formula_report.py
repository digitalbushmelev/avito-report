# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def main() -> None:
    env = os.environ.copy()
    env.update(
        {
            "REPORT_AVITO_ENV_FILE": "avito_second.env",
            "REPORT_BITRIX_ENV_FILE": "bitrix_second.env",
            "REPORT_AVITO_DATA_DIR": "data/avito_formula",
            "REPORT_AVITO_RAW_DIR": "data/avito_formula/raw",
            "REPORT_AVITO_CAMPAIGN_REGISTRY_FILE": "avito_formula_campaigns.json",
            "REPORT_OUTPUT_FILE": "formula_avito_report.html",
            "REPORT_INDEX_FILE": "formula/index.html",
            "REPORT_BRAND_NAME": "Авито Реклама · Формула",
            "REPORT_HEADING": "Сводка Авито Реклама · Формула",
            "REPORT_EYEBROW": "Внутренний отчет · Формула",
            "AVITO_CAMPAIGN_ID": "",
            "AVITO_CAMPAIGN_NAME": "",
            "AVITO_CAMPAIGN_IDS": "",
            "AVITO_CAMPAIGN_NAMES": "",
            "AVITO_CAMPAIGN_SLUGS": "",
            "BITRIX_ENTITY_TYPE": "deal",
            "BITRIX_UTM_SOURCE_FIELD": "UTM_SOURCE",
            "BITRIX_UTM_MEDIUM_FIELD": "UTM_MEDIUM",
            "BITRIX_UTM_CAMPAIGN_FIELD": "UTM_CAMPAIGN",
            "BITRIX_UTM_GROUPS_JSON": '[{"source":"avito","medium":"cpc"}]',
            "BITRIX_DEAL_CATEGORY_NAME": "Льготная ипотека",
            "BITRIX_STAGE_FIELD": "STAGE_ID",
            "BITRIX_COMMENT_FIELD": "COMMENTS",
        }
    )

    subprocess.run([sys.executable, str(ROOT / "update_report.py"), *sys.argv[1:]], cwd=ROOT, env=env, check=True)


if __name__ == "__main__":
    main()
