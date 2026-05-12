# MGT267 Watchdog

This project logs in to the Supply Chain Game, tracks cash, day, WH1 warehouse inventory, and overall team standings, then sends scheduled email reports.

## Editable Settings

Most settings are in `monitor_config.json`:

- `crawl.interval_minutes`: local watch interval, currently `60`.
- `monitor.target_team`: team to lock for comparison, currently `group7`.
- `monitor.warehouse_inventory_threshold`: inventory alert threshold, currently `450`.
- `monitor.warning_minutes`: warning email interval label, currently `15`.
- `monitor.send_report_every_run`: send one email each scheduled run.
- `monitor.report_min_interval_minutes`: minimum time between scheduled hourly report emails, currently `50`.
- `monitor.alert_rules`: editable watchlist rules for stockout risk, lost demand, days of cover, shipment coverage, and cash lead.
- `monitor.metric_thresholds`: optional min/max alert thresholds for the hourly report's warehouse, factory, and headquarters metrics.
- `ai.enabled`, `ai.model`, `ai.api_key_env`: Gemini recommendation settings.
- `email.recipients`: notification recipient list.
- `email.attach_excel`: attach the full data workbook to report emails.
- `email.footer`: email footer text and URL.
- `excel.exponential_smoothing_alpha`: alpha used by the Excel EMA formulas, currently `0.3`.
- `crawl.plot_sources`: warehouse, factory, and headquarters plot URLs included in hourly reports.

Configured recipients:

- `950154@gmail.com`

GitHub Actions schedules:

```yaml
# .github/workflows/warning.yml
- cron: "7,22,37,52 * * * *"

# .github/workflows/monitor.yml
- cron: "11,41 * * * *"

# .github/workflows/heartbeat.yml
repository_dispatch: watchdog-heartbeat

# .github/workflows/manual-export.yml
workflow_dispatch only
```

GitHub requires cron schedules to live in workflow files, so edit those lines if the cloud schedule needs to change.
The hourly workflow attempts twice per hour as a backup; `monitor.report_min_interval_minutes` throttles actual report emails so normal delivery remains about once per hour.
The minutes intentionally avoid exact hour and quarter-hour boundaries because GitHub scheduled workflows can be delayed or dropped during high-load times.
The heartbeat workflow starts immediately with `[heartbeat-now]`, checks warning rules, runs the hourly report throttle, waits 15 minutes, then dispatches the next heartbeat run.
Push with `[hourly-now]` or `[warning-now]` starts the corresponding workflow immediately.
The warning workflow checks every 15 minutes, but only sends email when a `monitor.alert_rules` entry with the `warning` channel is in `ALERT`.

## Manual Excel Export

To download the latest full crawler workbook without sending email or changing the running schedule:

1. Open GitHub `Actions`.
2. Select `Manual Data Export`.
3. Click `Run workflow`.
4. Open the completed run.
5. Download the `latest-crawl-data` artifact.

The artifact includes `supply_chain_data_latest.xlsx` plus the latest JSON/CSV outputs. This workflow does not restore or save the scheduled monitor cache.

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

Primary watchlist thresholds are in `monitor.alert_rules`. Each rule has:

- `enabled`: set to `false` to disable a rule.
- `metric`: source metric or derived metric.
- `operator`: one of `>`, `>=`, `<`, `<=`, `=`.
- `threshold`: editable numeric threshold.
- `severity`: `warning` or `critical`.
- `channels`: `hourly`, `warning`, or both.

The 15-minute workflow only sends when a rule with the `warning` channel is in `ALERT`.

Other legacy hourly report alert thresholds are in `monitor.metric_thresholds`. Set `min` or `max` to a number; leave unused thresholds as `null`.

Available metric keys:

- `warehouse_inventory:mail`
- `warehouse_inventory:warehouse`
- `warehouse_inventory:truck`
- `warehouse_shipments:Calopeia`
- `factory_wip:Calopeia`
- `hq_demand:Calopeia`
- `hq_lost_demand:Calopeia`
- `hq_cash_balance:value`
- `derived:days_of_cover`
- `derived:lost_demand_rate`
- `derived:shipment_to_demand_ratio`
- `derived:wip_to_demand_ratio`
- `derived:cash_lead_percent_vs_nearest`

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
- a `Watchlist` tab and EMA / Delta vs EMA formula columns on each Data tab

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
- `email_delivery_latest.json`

Email delivery is validated recipient-by-recipient. The workflow fails if SMTP does not accept every configured recipient, and `email_delivery_latest.json` records accepted/rejected recipients and message IDs.
