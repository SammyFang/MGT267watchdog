# MGT267 Watchdog

This project logs in to the Supply Chain Game, tracks cash, day, WH1 warehouse inventory, and overall team standings, then sends an hourly email report.

## Editable Settings

Most settings are in `monitor_config.json`:

- `crawl.interval_minutes`: local watch interval, currently `60`.
- `monitor.target_team`: team to lock for comparison, currently `group7`.
- `monitor.warehouse_inventory_threshold`: inventory alert threshold, currently `450`.
- `monitor.send_report_every_run`: send one email each scheduled run.
- `email.recipients`: notification recipient list.

GitHub Actions schedule is in `.github/workflows/monitor.yml`:

```yaml
- cron: "0 * * * *"
```

GitHub requires cron schedules to live in the workflow file, so edit that line if the cloud schedule needs to change.

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

Continuous local loop:

```powershell
npm run watch
```

Output files are written to `.monitor-state/`:

- `latest.json`
- `history.csv`
- `warehouse_inventory_latest.csv`
- `standing_gaps_latest.csv`
