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
- `game_rules`: Supply Chain Game economics used by Action Notes, including day `1460` obsolescence, 24-hour order loss, transport costs, capacity lead time, and the no-effect priority rule.
- `ai.enabled`, `ai.model`, `ai.api_key_env`: Gemini recommendation settings.
- `auto_adjust`: research-only adjustment planner settings. It writes suggested policy changes to email, JSON/CSV, and Excel, but never submits game forms.
- `backtest`: per-trigger backtest settings. It uses the plot data crawled during the current GitHub Actions run and has no local file dependency.
- `email.recipients`: notification recipient list.
- `email.game_entry_url`: clickable Supply Chain Game entry link included in every email.
- `email.attach_excel`: attach the full data workbook to report emails.
- `email.footer`: email footer text and URL.
- `excel.exponential_smoothing_alpha`: alpha used by the Excel EMA formulas, currently `0.3`.
- `crawl.plot_sources`: warehouse, factory, and headquarters plot URLs included in hourly reports.
- `crawl.policy_pages`: headquarters, Calopeia factory, and Calopeia/Sorange/Tyran/Entworpe warehouse policy pages included in reports and Excel.

Configured recipients:

- `950154@gmail.com`
- `wgong009@ucr.edu`
- `hhuan238@ucr.edu`
- `yfang097@ucr.edu`

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

The artifact includes `supply_chain_data_latest.xlsx` plus the latest JSON/CSV outputs, including `policy_snapshot_latest.csv`, the research-only `adjustment_plan_latest` files, and the per-trigger `backtest_latest` outputs. This workflow does not restore or save the scheduled monitor cache.

## Standing Gap Formula

The report locks on `monitor.target_team` and calculates every row against that target:

```text
gap_amount = target_cash - team_cash
gap_percent = gap_amount / team_cash * 100
```

Positive values mean `group7` is ahead of that team. Negative values mean `group7` is behind that team.

When WH1 warehouse inventory is at or above `monitor.warehouse_inventory_threshold`, the hourly email subject is prefixed with `ALERT`.
That threshold is treated as a review checkpoint. Action Notes compare inventory with the demand of all regions currently served by the Calopeia warehouse before suggesting any inventory reduction.

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
- `derived:calopeia_served_demand`
- `derived:calopeia_served_lost_demand`
- `derived:calopeia_served_shipments`
- `derived:calopeia_served_region_count`
- `derived:lost_demand_rate`
- `derived:shipment_to_demand_ratio`
- `derived:wip_to_demand_ratio`
- `derived:cash_lead_percent_vs_nearest`

## Auto-Adjustment Research Plan

`auto_adjust` is intentionally research-only. The code does not submit Factory or Warehouse forms, does not POST policy changes, and does not contain an apply mode. It uses scraped current policy values as the baseline when available, then falls back to `auto_adjust.policy_baseline`.

Editable fields:

- `auto_adjust.targets`: desired bands for days of cover, inventory, lost demand, shipment coverage, and WIP coverage.
- `auto_adjust.max_change_per_run`: maximum suggested change per crawler run for `order_point` and `quantity`.
- `auto_adjust.bounds`: min/max allowed values used when calculating suggested values.
- `auto_adjust.policy_baseline`: current Factory and Warehouse policy values used as the baseline for suggested changes.

The current days-of-cover research band is `2` to `5` days with target `3`, calculated from the total demand of regions served by Calopeia. This matches the high-cover alert at `> 5` days and the game rule that inventory becomes worthless on day `1460`.
`priority1` is still captured in policy tables for completeness, but it is intentionally excluded from recommendations because the assignment states priority level has no effect.

Outputs:

- Email section: `Auto-Adjustment Research Plan`
- Excel tab: `Adjustment Plan`
- JSON: `.monitor-state/adjustment_plan_latest.json`
- CSV: `.monitor-state/adjustment_plan_latest.csv`

## Per-Trigger Backtest

Every `monitor`, `warning-email`, `test-email`, and manual export trigger runs a fresh backtest from the plot data crawled during that same run. It does not rely on local `.monitor-state` files. Restored GitHub cache is only used for report throttling and previous-run deltas, not as the backtest data source.

Editable fields are in `backtest`:

- `inventory_threshold_candidates`: warehouse inventory thresholds to compare.
- `days_of_cover_high_candidates`: high-cover thresholds to compare.
- `days_of_cover_low_candidates`: low-cover thresholds to compare against future lost demand.
- `shipment_to_demand_ratio_candidates`: shipment coverage thresholds to compare against future lost demand.
- `horizon_days`: future window for shortage-event scoring.
- `excess_cover_days`: high-cover event definition.
- `lost_demand_threshold`: lost-demand event definition.

Outputs:

- Email section: `Backtest Snapshot`
- Excel tabs: `Backtest`, `Backtest Summary`, `Backtest Thresholds`, `Backtest Daily`
- JSON: `.monitor-state/backtest_latest.json`
- CSV: `.monitor-state/backtest_summary_latest.csv`
- CSV: `.monitor-state/backtest_daily_latest.csv`
- XLSX: `.monitor-state/backtest_report_latest.xlsx`

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

Run only a fresh crawler backtest without sending mail:

```powershell
$env:EMAIL_ENABLED='false'; $env:EMAIL_DRY_RUN='1'; npm run backtest
```

Hourly reports include:

- warehouse inventory and shipments
- factory WIP
- headquarters demand, lost demand, and cash balance
- current source values plus 1-hour change and 1-hour change rate when a previous hourly state exists
- current headquarters/factory/warehouse policy values from Calopeia plus Sorange/Tyran/Entworpe warehouse policy pages
- one `.xlsx` attachment with every scraped Data table in separate tabs, plus summary, standing, and policy tabs
- a `Watchlist` tab and EMA / Delta vs EMA formula columns on each Data tab
- an `Adjustment Plan` tab with research-only suggested policy changes and safety flags
- a dedicated `Backtest` tab plus detailed backtest tabs generated from the same trigger's crawled plot data

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
- `policy_snapshot_latest.csv`
- `adjustment_plan_latest.json`
- `adjustment_plan_latest.csv`
- `backtest_latest.json`
- `backtest_summary_latest.csv`
- `backtest_daily_latest.csv`
- `backtest_report_latest.xlsx`
- `supply_chain_data_latest.xlsx`
- `email_delivery_latest.json`

Email delivery is validated recipient-by-recipient. The workflow fails if SMTP does not accept every configured recipient, and `email_delivery_latest.json` records accepted/rejected recipients and message IDs.
