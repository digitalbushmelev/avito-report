# Отчет по Авито Рекламе

`index.html` — основной HTML-отчет для GitHub Pages по кампании `Дружеский`.
`avito_report.html` генерируется как локальная копия с тем же содержимым.

Данные собираются из двух источников:

- Avito Реклама API: показы, клики, расход, CTR, CPC.
- Bitrix: сделки по UTM-меткам, лиды и `клик -> лид`.

## Быстрый запуск

```powershell
python update_report.py --skip-avito
```

Команда не дергает Avito API, а только пересобирает отчет по имеющимся данным и Bitrix.

## Avito API

1. Скопируйте шаблон:

```powershell
Copy-Item avito.env.example avito.env
```

2. Заполните в `avito.env`:

- `AVITO_CLIENT_ID`;
- `AVITO_CLIENT_SECRET`;
- `AVITO_ACCOUNT_ID`;
- `AVITO_CAMPAIGN_ID` или `AVITO_CAMPAIGN_IDS`.

3. Включите API:

```env
AVITO_ENABLE_API=1
```

По умолчанию используется стандартный endpoint статистики кампании из Avito Ads API:

```text
POST /ads/v1/account/{accountID}/campaigns/{campaignID}/stats
```

Для нескольких кампаний используйте:

```env
AVITO_CAMPAIGN_IDS=843568278,123456789
AVITO_CAMPAIGN_NAMES=Дружеский,Новая кампания
AVITO_CAMPAIGN_SLUGS=druzheskiy,new-campaign
```

Если `AVITO_AUTO_DISCOVER_CAMPAIGNS=1`, скрипт сначала пробует получить список кампаний аккаунта через `GET /ads/v1/account/{account_id}/campaigns`. Если Avito не отдает список для текущего ключа, используется явный список `AVITO_CAMPAIGN_IDS` или старый `AVITO_CAMPAIGN_ID`.

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
- `AVITO_CAMPAIGN_ID` или `AVITO_CAMPAIGN_IDS`;
- `AVITO_CAMPAIGN_NAMES` и `AVITO_CAMPAIGN_SLUGS` — опционально для нескольких кампаний;
- `AVITO_AUTO_DISCOVER_CAMPAIGNS` — опционально, по умолчанию включено;
- `AVITO_CAMPAIGN_LIST_PATH_TEMPLATE` — опционально, если Avito изменит endpoint списка кампаний;
- `BITRIX_WEBHOOK_URL`;
- `BITRIX_DATE_FIELD`;
- `BITRIX_CAMPAIGN_DATE_MAP_JSON` — опционально, если нужно раскладывать сделки по Avito-кампаниям через дату создания;
- `BITRIX_UNTOUCHED_STAGE_NAMES` — опционально, стадии обращений без обработки;
- `BITRIX_DEAL_CATEGORY_ID` — необязательно, если название воронки резолвится через API;
- `REPORT_START_DATE`.

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

Если в отчете несколько Avito-кампаний, сделки можно привязать к конкретной строке кампании через дату создания. Текущий принцип: до `14.06.2026` включительно — первая кампания, с `15.06.2026` — расширенная.

```json
{
  "843568278": {
    "date_to": "2026-06-14"
  },
  "460053704": {
    "date_from": "2026-06-15"
  }
}
```

Без этого маппинга сделки остаются в общем итоге и в разбивке по `utm_campaign`, но не приписываются к конкретной Avito-кампании.

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
