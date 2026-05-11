# MGT267 Watchdog

This project logs in to the Supply Chain Game, tracks cash, day, WH1 warehouse inventory, and overall team standings, then sends scheduled email reports.

## Editable Settings

Most settings are in `monitor_config.json`:

- `crawl.interval_minutes`: local watch interval, currently `60`.
- `monitor.target_team`: team to lock for comparison, currently `group7`.
- `monitor.warehouse_inventory_threshold`: inventory alert threshold, currently `450`.
- `monitor.warning_minutes`: warning email interval label, currently `15`.
- `monitor.send_report_every_run`: send one email each scheduled run.
- `monitor.metric_thresholds`: optional min/max alert thresholds for the hourly report's warehouse, factory, and headquarters metrics.
- `ai.enabled`, `ai.model`, `ai.api_key_env`: Gemini recommendation settings.
- `email.recipients`: notification recipient list.
- `email.attach_excel`: attach the full data workbook to report emails.
- `email.footer`: email footer text and URL.
- `crawl.plot_sources`: warehouse, factory, and headquarters plot URLs included in hourly reports.

Configured recipients:

- `950154@gmail.com`

GitHub Actions schedules:

```yaml
# .github/workflows/warning.yml
- cron: "7,22,37,52 * * * *"

# .github/workflows/monitor.yml
- cron: "11 * * * *"
```

GitHub requires cron schedules to live in workflow files, so edit those lines if the cloud schedule needs to change.
The minutes intentionally avoid exact hour and quarter-hour boundaries because GitHub scheduled workflows can be delayed or dropped during high-load times.
Push with `[hourly-now]` or `[warning-now]` starts the corresponding workflow immediately.
The warning workflow checks every 15 minutes, but only sends email when WH1 warehouse inventory is at or above the configured threshold.

## Standing Gap Formula

The report locks on `monitor.target_team` and calculates every row against that target:

```text
gap_amount = target_cash - team_cash
gap_percent = gap_amount / team_cash * 100
```

Positive values mean `group7` is ahead of that team. Negative values mean `group7` is behind that team.

When WH1 warehouse inventory is at or above `monitor.warehouse_inventory_threshold`, the hourly email subject is prefixed with `ALERT`.

## Required Secrets

Set these in GitHub repo settings: `Settings -> Secrets and variables -> Actions -> New repository secret`.

- `SC_TEAM_ID`
- `SC_PASSWORD`
- `SC_INSTITUTION`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_SECURE`
- `SMTP_USER`
- `SMTP_PASSWORD`
- `SMTP_FROM`

For Gmail SMTP, use an app password, not the normal Gmail login password.

Typical Gmail SMTP values:

```text
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_SECURE=true
SMTP_USER=your-email@gmail.com
SMTP_PASSWORD=your-app-password
SMTP_FROM=your-email@gmail.com
```

## Optional AI Secret

Gemini is optional but recommended for action notes in the email. Create a Gemini API key in Google AI Studio, then add it to GitHub as:

```text
Name: GEMINI_API_KEY
Value: your Gemini API key
```

GitHub path: `Settings -> Secrets and variables -> Actions -> New repository secret`.

## Editable Alert Thresholds

The 15-minute warning email threshold is:

```json
"warehouse_inventory_threshold": 450
```

Other hourly report alert thresholds are in `monitor.metric_thresholds`. Set `min` or `max` to a number; leave unused thresholds as `null`.

Available metric keys:

- `warehouse_inventory:mail`
- `warehouse_inventory:truck`
- `warehouse_shipments:Calopeia`
- `factory_wip:Calopeia`
- `hq_demand:Calopeia`
- `hq_lost_demand:Calopeia`
- `hq_cash_balance:value`

## Run Locally

```powershell
copy .env.example .env
# edit .env first
npm install
npm run monitor
```

Local test without sending mail:

```powershell
$env:EMAIL_DRY_RUN='1'; npm run monitor
```

Send the email templates without sending mail locally:

```powershell
$env:EMAIL_DRY_RUN='1'; npm run warning-email
$env:EMAIL_DRY_RUN='1'; npm run test-email
```

Hourly reports include:

- warehouse inventory and shipments
- factory WIP
- headquarters demand, lost demand, and cash balance
- current source values plus 1-hour change and 1-hour change rate when a previous hourly state exists
- one `.xlsx` attachment with every scraped Data table in separate tabs, plus summary and standing tabs

The `Email Smoke Test` GitHub workflow sends one real test email with repository secrets:

- `[TEST] hourly report`

It can be run manually from GitHub Actions. It also runs on push only when the commit message contains `[email-test]`.

Continuous local loop:

```powershell
npm run watch
```

Output files are written to `.monitor-state/`:

- `latest.json`
- `history.csv`
- `warehouse_inventory_latest.csv`
- `standing_gaps_latest.csv`
- `operational_snapshot_latest.csv`
- `supply_chain_data_latest.xlsx`
