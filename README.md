# MGT267 Watchdog

This project logs in to the Supply Chain Game, tracks cash, day, and WH1 warehouse inventory, and sends an email when warehouse inventory reaches the configured threshold.

## Editable Settings

Most settings are in `monitor_config.json`:

- `crawl.interval_minutes`: local watch interval.
- `monitor.warehouse_inventory_threshold`: alert threshold, currently `450`.
- `monitor.notification_cooldown_minutes`: `0` means alert once per crossing from below threshold to at/above threshold.
- `email.recipients`: notification recipient list.

GitHub Actions schedule is in `.github/workflows/monitor.yml`:

```yaml
- cron: "*/15 * * * *"
```

GitHub requires cron schedules to live in the workflow file, so edit that line if the cloud schedule needs to change.

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

## Run Locally

```powershell
copy .env.example .env
# edit .env first
npm install
npm run monitor
```

Continuous local loop:

```powershell
npm run watch
```

Output files are written to `.monitor-state/`:

- `latest.json`
- `history.csv`
- `warehouse_inventory_latest.csv`
