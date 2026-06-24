# Отчет по Авито Рекламе

`index.html` — основной HTML-отчет для GitHub Pages по кампании `Дружеский`.
`avito_report.html` генерируется как локальная копия с тем же содержимым.
`formula/index.html` и `formula_avito_report.html` — отдельная версия отчета для кабинета Формулы.

Данные собираются из двух источников:

- Avito Реклама API: показы, клики, расход, CTR, CPC.
- Bitrix: сделки по UTM-меткам, лиды и `клик -> лид`.

## Быстрый запуск

```powershell
python update_report.py --skip-avito
```

Команда не дергает Avito API, а только пересобирает отчет по имеющимся данным и Bitrix.

Для второго кабинета Формулы:

```powershell
python update_formula_report.py --skip-avito
```

Принудительно обновить кампании и статистику Формулы из Avito API:

```powershell
python update_formula_report.py --force
```

Этот запуск использует `avito_second.env`, `bitrix_second.env`, пишет кэш в `data/avito_formula`, реестр кампаний в `avito_formula_campaigns.json`, а HTML — в `formula/index.html` и `formula_avito_report.html`.

## Avito API

1. Скопируйте шаблон:

```powershell
Copy-Item avito.env.example avito.env
```

2. Заполните в `avito.env`:

- `AVITO_CLIENT_ID`;
- `AVITO_CLIENT_SECRET`;
- `AVITO_ACCOUNT_ID`.

3. Включите API:

```env
AVITO_ENABLE_API=1
```

По умолчанию используется стандартный endpoint статистики кампании из Avito Ads API:

```text
POST /ads/v1/account/{accountID}/campaigns/{campaignID}/stats
```

Кампании задаются в `avito_campaigns.json`. Этот файл коммитится в репозиторий и служит fallback-реестром, если Avito API не отдает список кампаний:

```json
{
  "campaigns": [
    {"id": "843568278", "name": "Дружеский", "slug": "druzheskiy", "lead_date_to": "2026-06-14"},
    {"id": "460053704", "name": "Дружеский - расширенная", "slug": "druzheskiy-rasshirennaya", "lead_date_from": "2026-06-15", "lead_date_to": "2026-06-21"},
    {"id": "771269910", "name": "Эклипт", "slug": "eklipt", "lead_date_from": "2026-06-22"}
  ]
}
```

Если нужно переопределить кампании через переменные окружения, используйте:

```env
AVITO_CAMPAIGN_IDS=843568278,123456789
AVITO_CAMPAIGN_NAMES=Дружеский,Новая кампания
AVITO_CAMPAIGN_SLUGS=druzheskiy,new-campaign
```

Если `AVITO_AUTO_DISCOVER_CAMPAIGNS=1`, скрипт сначала пробует получить список кампаний аккаунта через `GET /ads/v1/account/{account_id}/campaigns`. Если Avito не отдает список для текущего ключа, используется `avito_campaigns.json`, затем явный список `AVITO_CAMPAIGN_IDS` или старый `AVITO_CAMPAIGN_ID`.
Если список кампаний недоступен, скрипт дополнительно ищет campaignID через список креативов `POST /ads/v1/account/{account_id}/creatives` и сохраняет найденные `utm_campaign` как `lead_keys` для атрибуции сделок.
При успешном автообнаружении и `AVITO_WRITE_DISCOVERED_CAMPAIGNS=1` скрипт обновляет `avito_campaigns.json`, а GitHub Actions коммитит этот файл вместе с отчетом.

`update_report.py` сохраняет нормализованные данные по каждой кампании в `data/avito/*.json`. Если файл кампании уже обновлялся сегодня, ручной запуск без `--force` не тратит баллы Avito API. Для принудительного обновления:

```powershell
python update_report.py --force
```

## Обновление в GitHub Actions

В репозитории подготовлен workflow:

```text
.github/workflows/update-report.yml
```

Он запускается каждый день в `03:00`, `08:00` и `11:00 UTC`, то есть в `08:00`, `13:00` и `16:00` по Екатеринбургу, принудительно обновляет Avito, забирает сделки из Bitrix, генерирует `index.html` и коммитит результат через `GITHUB_TOKEN`. Локальная авторизация GitHub на этой машине не используется.

Ручной запуск workflow без `force_avito=true` пересобирает отчет без лишнего Avito-запроса, если кеш уже свежий.

Для работы workflow заведите GitHub Secrets:

- `AVITO_ENABLE_API`;
- `AVITO_CLIENT_ID`;
- `AVITO_CLIENT_SECRET`;
- `AVITO_ACCOUNT_ID`;
- `AVITO_CAMPAIGN_ID`, `AVITO_CAMPAIGN_IDS`, `AVITO_CAMPAIGN_NAMES`, `AVITO_CAMPAIGN_SLUGS` — опционально, если нужно переопределить `avito_campaigns.json`;
- `AVITO_AUTO_DISCOVER_CAMPAIGNS` — опционально, по умолчанию включено;
- `AVITO_CAMPAIGN_LIST_PATH_TEMPLATE` — опционально, если Avito изменит endpoint списка кампаний;
- `BITRIX_WEBHOOK_URL`;
- `BITRIX_DATE_FIELD`;
- `BITRIX_CAMPAIGN_DATE_MAP_JSON` — опционально, если нужно переопределить даты из `avito_campaigns.json`;
- `BITRIX_UNTOUCHED_STAGE_NAMES` — опционально, стадии обращений без обработки;
- `BITRIX_DEAL_CATEGORY_ID` — необязательно, если название воронки резолвится через API;
- `REPORT_START_DATE`.

Для отчета Формулы workflow использует отдельные секреты. Если они не заполнены, шаг Формулы пропускается:

- `AVITO_SECOND_ENABLE_API`;
- `AVITO_SECOND_CLIENT_ID`;
- `AVITO_SECOND_CLIENT_SECRET`;
- `AVITO_SECOND_ACCOUNT_ID`;
- `AVITO_SECOND_TOKEN_URL`, `AVITO_SECOND_BASE_URL` — опционально;
- `AVITO_SECOND_STATS_METHOD`, `AVITO_SECOND_STATS_PATH_TEMPLATE`, `AVITO_SECOND_STATS_BODY_JSON` — опционально;
- `AVITO_SECOND_GROUP_BY`, `AVITO_SECOND_LOOKBACK_DAYS` — опционально;
- `BITRIX_SECOND_WEBHOOK_URL`;
- `BITRIX_SECOND_DATE_FIELD` — опционально;
- `REPORT_SECOND_START_DATE` — опционально.

Для GitHub Pages включите публикацию из ветки, где лежит `index.html`, с корнем `/`.

В отчете есть фильтр периода по датам. Он работает на стороне браузера: меняет плашку `Период`, KPI, динамику, таблицу кампаний, воронку, прогноз и топы без повторной генерации HTML.
Границы фильтра берутся из фактических дат Avito/Bitrix-данных. `REPORT_START_DATE` нужен только как нижняя граница запроса сделок в Bitrix.

## Bitrix

Скрипт читает `bitrix.env` и ищет сделки через `crm.deal.list`.

Фильтры:

```text
Воронка = Комфорт: прямые продажи

Группа 1:
  utm_source = avito_media
  utm_medium = banner

Группа 2:
  utm_source = avito_reklama
  utm_medium = cpc
```

В REST API Bitrix эти поля запрашиваются как `UTM_SOURCE`, `UTM_MEDIUM`, `UTM_CAMPAIGN`. Сделка попадает в отчет только по паре `utm_source + utm_medium`; `utm_campaign` не фильтрует сделки и сохраняется отдельно для разбивки по креативам.

Если в отчете несколько Avito-кампаний, сделки привязываются к конкретной строке кампании через дату создания. Основной источник этих границ — поля `lead_date_from` и `lead_date_to` в `avito_campaigns.json`. Если Avito API списка кампаний доступен и появляется новая кампания, скрипт выставляет ей `lead_date_from` по первой активной дате из статистики и закрывает предыдущую открытую кампанию днем раньше.

Текущий принцип:

- до `14.06.2026` включительно — `Дружеский`;
- `15.06.2026`–`21.06.2026` — `Дружеский - расширенная`;
- с `22.06.2026` — `Эклипт`.

Для ручного переопределения можно использовать `BITRIX_CAMPAIGN_DATE_MAP_JSON`:

```json
{
  "843568278": {
    "date_to": "2026-06-14"
  },
  "460053704": {
    "date_from": "2026-06-15",
    "date_to": "2026-06-21"
  },
  "771269910": {
    "date_from": "2026-06-22"
  }
}
```

Если нет ни `avito_campaigns.json`, ни date-map, сделки остаются в общем итоге и в разбивке по `utm_campaign`, но не приписываются к конкретной Avito-кампании.

Не учитываются:

- сделки, где в `COMMENTS` есть телефон `71111111111`;
- сделки в стадиях `Дубль. Создана новая сделка` и `Наш сотрудник`.

В верхней сводке `Лиды` — это сделки в качественных стадиях. `Обращения` — проваленные сделки и сделки в стадиях без обработки. По умолчанию к обращениям относятся стадии `ЛИДГЕН` и `Совершить первый контакт`; список можно переопределить через `BITRIX_UNTOUCHED_STAGE_NAMES`.

По умолчанию скрипт сам получает `CATEGORY_ID` по названию воронки. Если в Bitrix API это не сработает, укажите ID в `BITRIX_DEAL_CATEGORY_ID`.

Если в `bitrix.env` остался `BITRIX_ENTITY_TYPE=lead`, генератор все равно использует сделки для этого отчета.

## Файлы

- `generate_avito_report.py` — строит HTML.
- `update_report.py` — обновляет Avito-кэш и пересобирает отчет.
- `avito.env.example` — шаблон для Avito API.
- `bitrix_config.example.json` — запасной пример Bitrix-настроек.
- `.github/workflows/update-report.yml` — обновление по расписанию и коммит отчета.
