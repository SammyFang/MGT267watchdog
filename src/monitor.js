const fs = require("node:fs");
const path = require("node:path");

require("dotenv").config();

const nodemailer = require("nodemailer");

const CONFIG_PATH = path.resolve(process.cwd(), "monitor_config.json");

function readJson(filePath, fallback = null) {
  if (!fs.existsSync(filePath)) {
    return fallback;
  }

  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function requiredEnv(name) {
  const value = process.env[name];

  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }

  return value;
}

function optionalEnv(name, fallback = "") {
  return process.env[name] || fallback;
}

function collectCookies(headers, cookieJar) {
  const setCookie = headers.getSetCookie
    ? headers.getSetCookie()
    : headers.get("set-cookie")
      ? [headers.get("set-cookie")]
      : [];

  for (const cookie of setCookie) {
    const [pair] = cookie.split(";");
    if (!pair) continue;
    const [name, ...valueParts] = pair.split("=");
    if (!name || valueParts.length === 0) continue;
    cookieJar.set(name.trim(), valueParts.join("=").trim());
  }
}

function cookieHeader(cookieJar) {
  return [...cookieJar.entries()]
    .map(([name, value]) => `${name}=${value}`)
    .join("; ");
}

async function request(url, options, cookieJar) {
  const headers = new Headers(options.headers || {});
  const cookies = cookieHeader(cookieJar);

  if (cookies) {
    headers.set("cookie", cookies);
  }

  const response = await fetch(url, {
    ...options,
    headers,
    redirect: "manual",
  });

  collectCookies(response.headers, cookieJar);
  return response;
}

async function follow(response, cookieJar, limit = 5) {
  let current = response;

  for (let i = 0; i < limit; i += 1) {
    if (![301, 302, 303, 307, 308].includes(current.status)) {
      return current;
    }

    const location = current.headers.get("location");
    if (!location) {
      return current;
    }

    current = await request(
      new URL(location, current.url).toString(),
      { method: "GET" },
      cookieJar,
    );
  }

  throw new Error("Too many redirects during login");
}

function decodeHtml(value) {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

function isLoggedIn(html, teamId) {
  return (
    html.includes(`Name: <b>${teamId}</b>`) ||
    (html.includes("SCQuit") && html.includes("Supply Chain Game"))
  );
}

function extractDashboardValues(html) {
  const cashMatch = html.match(/Cash:\s*<b>([^<]+)<\/b>/i);
  const dayMatch = html.match(/Day:\s*<b>([^<]+)<\/b>/i);

  if (cashMatch && dayMatch) {
    return {
      cash: decodeHtml(cashMatch[1]),
      day: decodeHtml(dayMatch[1]),
    };
  }

  const boldValues = [...html.matchAll(/<b>([^<]+)<\/b>/gi)].map((match) =>
    decodeHtml(match[1]),
  );

  if (boldValues.length >= 3) {
    return {
      cash: boldValues[1],
      day: boldValues[2],
    };
  }

  throw new Error("Logged in, but could not find cash/day values");
}

function extractCashNumber(cashText) {
  const normalized = cashText.replace(/[$,]/g, "");
  const value = Number.parseFloat(normalized);
  return Number.isFinite(value) ? value : null;
}

function extractPlotLines(html) {
  const linePattern =
    /\{label:\s*'([^']+)'\s*,\s*name:\s*'([^']+)'\s*,\s*points:\s*'([^']*)'/g;
  const lines = [];

  for (const match of html.matchAll(linePattern)) {
    const points = match[3]
      .trim()
      .split(/\s+/)
      .filter(Boolean)
      .map(Number);

    if (points.length % 2 !== 0 || points.some((point) => Number.isNaN(point))) {
      throw new Error(`Invalid points data for ${match[1]}`);
    }

    lines.push({
      label: decodeHtml(match[1]),
      name: match[2],
      points,
    });
  }

  if (lines.length === 0) {
    throw new Error("Could not find warehouse inventory plot data");
  }

  return lines;
}

function extractAxisLabel(html) {
  const match = html.match(/hAxisLabel:\s*'([^']+)'/i);
  return match ? decodeHtml(match[1]) : "day";
}

function formatNumber(value, precision, groupingSeparator = ",", groupingSize = 3) {
  if (!Number.isFinite(value)) {
    return "";
  }

  const [integerPart, decimalPart] = value.toFixed(precision).split(".");
  const sign = integerPart.startsWith("-") ? "-" : "";
  const digits = sign ? integerPart.slice(1) : integerPart;
  const groups = [];

  for (let i = digits.length; i > 0; i -= groupingSize) {
    groups.unshift(digits.slice(Math.max(0, i - groupingSize), i));
  }

  return `${sign}${groups.join(groupingSeparator)}${
    decimalPart ? `.${decimalPart}` : ""
  }`;
}

function csvValue(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function buildPlotRows(lines) {
  const indexes = Array(lines.length).fill(0);
  const totalPoints = lines.reduce((sum, line) => sum + line.points.length / 2, 0);
  const rows = [];
  let consumed = 0;

  while (consumed < totalPoints) {
    let minX = Number.MAX_VALUE;
    let selected = [];

    for (let i = 0; i < lines.length; i += 1) {
      const pointIndex = indexes[i];
      const x = lines[i].points[pointIndex];

      if (x === undefined) {
        continue;
      }

      if (x < minX) {
        minX = x;
        selected = [i];
      } else if (x === minX) {
        selected.push(i);
      }
    }

    if (selected.length === 0) {
      break;
    }

    const rawRow = Array(lines.length + 1).fill(null);
    const formattedRow = Array(lines.length + 1).fill("");
    rawRow[0] = minX;
    formattedRow[0] = formatNumber(minX, 3);

    for (const lineIndex of selected) {
      const value = lines[lineIndex].points[indexes[lineIndex] + 1];
      rawRow[lineIndex + 1] = value;
      formattedRow[lineIndex + 1] = formatNumber(value, 0);
      indexes[lineIndex] += 2;
    }

    rows.push({ raw: rawRow, formatted: formattedRow });
    consumed += selected.length;
  }

  return rows;
}

function parseWarehouseInventoryTable(html) {
  const lines = extractPlotLines(html);
  const header = [extractAxisLabel(html), ...lines.map((line) => line.label)];
  const rows = buildPlotRows(lines);
  const warehouseIndex = header.indexOf("warehouse");

  if (warehouseIndex < 0) {
    throw new Error("Could not find warehouse column in inventory data");
  }

  let latestWarehouse = null;

  for (const row of rows) {
    const value = row.raw[warehouseIndex];

    if (value !== null && Number.isFinite(value)) {
      latestWarehouse = {
        day: row.formatted[0],
        dayNumber: row.raw[0],
        inventory: value,
      };
    }
  }

  if (!latestWarehouse) {
    throw new Error("Could not find latest warehouse inventory value");
  }

  return { header, rows, latestWarehouse };
}

function buildWarehouseCsv({ header, rows }) {
  const csvLines = [
    header.map(csvValue).join(","),
    ...rows.map((row) => row.formatted.map(csvValue).join(",")),
  ];

  return `\uFEFF${csvLines.join("\r\n")}\r\n`;
}

function appendHistory(filePath, record) {
  const exists = fs.existsSync(filePath);
  const header = [
    "checked_at",
    "cash",
    "cash_number",
    "dashboard_day",
    "warehouse_inventory",
    "warehouse_day",
    "threshold",
    "notified",
  ];
  const row = [
    record.checkedAt,
    record.cash,
    record.cashNumber ?? "",
    record.dashboardDay,
    record.warehouseInventory,
    record.warehouseDay,
    record.threshold,
    record.notified ? "yes" : "no",
  ];
  const line = row.map(csvValue).join(",");

  fs.appendFileSync(
    filePath,
    `${exists ? "" : `${header.map(csvValue).join(",")}\n`}${line}\n`,
    "utf8",
  );
}

function localizedTime(isoString, timezone) {
  return new Intl.DateTimeFormat("en-US", {
    timeZone: timezone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(new Date(isoString));
}

function shouldNotify({ currentInventory, previousInventory, lastNotifiedAt, threshold, cooldownMinutes }) {
  if (currentInventory < threshold) {
    return false;
  }

  const crossedFromBelow =
    previousInventory === null ||
    previousInventory === undefined ||
    previousInventory < threshold;

  if (crossedFromBelow) {
    return true;
  }

  if (!cooldownMinutes || cooldownMinutes <= 0 || !lastNotifiedAt) {
    return false;
  }

  const elapsedMs = Date.now() - new Date(lastNotifiedAt).getTime();
  return elapsedMs >= cooldownMinutes * 60 * 1000;
}

function createTransport() {
  const host = requiredEnv("SMTP_HOST");
  const port = Number.parseInt(requiredEnv("SMTP_PORT"), 10);
  const secureText = optionalEnv("SMTP_SECURE", "false").toLowerCase();

  if (!Number.isFinite(port)) {
    throw new Error("SMTP_PORT must be a number");
  }

  return nodemailer.createTransport({
    host,
    port,
    secure: secureText === "true" || port === 465,
    auth: {
      user: requiredEnv("SMTP_USER"),
      pass: requiredEnv("SMTP_PASSWORD"),
    },
  });
}

async function sendAlert(config, record, previousInventory) {
  if (!config.email.enabled) {
    console.log("Email alert skipped because email.enabled=false");
    return;
  }

  const recipients = config.email.recipients || [];

  if (recipients.length === 0) {
    throw new Error("Email alert is enabled, but email.recipients is empty");
  }

  const transporter = createTransport();
  const subject = `${config.email.subject_prefix}: warehouse inventory ${record.warehouseInventory}`;
  const text = [
    "Warehouse inventory threshold reached.",
    "",
    `Checked at: ${record.checkedAtLocal} (${config.crawl.timezone})`,
    `Dashboard day: ${record.dashboardDay}`,
    `Cash: ${record.cash}`,
    `Warehouse inventory: ${record.warehouseInventory}`,
    `Warehouse inventory day: ${record.warehouseDay}`,
    `Threshold: ${record.threshold}`,
    `Previous warehouse inventory: ${previousInventory ?? "n/a"}`,
  ].join("\n");

  await transporter.sendMail({
    from: requiredEnv("SMTP_FROM"),
    to: recipients.join(","),
    subject,
    text,
  });
}

async function crawl(config) {
  const teamId = requiredEnv(config.credentials.team_id_env);
  const password = requiredEnv(config.credentials.password_env);
  const institution = optionalEnv(
    config.credentials.institution_env,
    config.credentials.default_institution,
  );
  const cookieJar = new Map();

  await request(config.crawl.entry_url, { method: "GET" }, cookieJar);

  const form = new URLSearchParams({
    id: teamId,
    password,
    institution,
    ismobile: "false",
  });

  const loginResponse = await follow(
    await request(
      config.crawl.login_url,
      {
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
        },
        body: form,
      },
      cookieJar,
    ),
    cookieJar,
  );
  const dashboardHtml = await loginResponse.text();

  if (!loginResponse.ok || !isLoggedIn(dashboardHtml, teamId)) {
    throw new Error(`Login failed with HTTP ${loginResponse.status}`);
  }

  const dashboard = extractDashboardValues(dashboardHtml);
  const inventoryResponse = await request(
    config.crawl.warehouse_inventory_url,
    { method: "GET" },
    cookieJar,
  );
  const inventoryHtml = await inventoryResponse.text();

  if (!inventoryResponse.ok) {
    throw new Error(`Warehouse inventory failed with HTTP ${inventoryResponse.status}`);
  }

  const inventoryTable = parseWarehouseInventoryTable(inventoryHtml);

  return {
    dashboard,
    inventoryTable,
  };
}

async function runOnce(config) {
  const statePath = path.resolve(process.cwd(), config.output.latest_json);
  const historyPath = path.resolve(process.cwd(), config.output.history_csv);
  const inventoryCsvPath = path.resolve(
    process.cwd(),
    config.output.warehouse_inventory_csv,
  );
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));

  const previousState = readJson(statePath, {});
  const { dashboard, inventoryTable } = await crawl(config);
  const threshold = Number(config.monitor.warehouse_inventory_threshold);
  const cooldownMinutes = Number(config.monitor.notification_cooldown_minutes || 0);
  const currentInventory = inventoryTable.latestWarehouse.inventory;
  const previousInventory = previousState.last_warehouse_inventory;
  const checkedAt = new Date().toISOString();
  let notified = false;

  const record = {
    checkedAt,
    checkedAtLocal: localizedTime(checkedAt, config.crawl.timezone),
    cash: dashboard.cash,
    cashNumber: extractCashNumber(dashboard.cash),
    dashboardDay: dashboard.day,
    warehouseInventory: currentInventory,
    warehouseDay: inventoryTable.latestWarehouse.day,
    threshold,
  };

  if (
    shouldNotify({
      currentInventory,
      previousInventory,
      lastNotifiedAt: previousState.last_notified_at,
      threshold,
      cooldownMinutes,
    })
  ) {
    await sendAlert(config, record, previousInventory);
    notified = true;
  }

  fs.writeFileSync(inventoryCsvPath, buildWarehouseCsv(inventoryTable), "utf8");
  appendHistory(historyPath, { ...record, notified });
  fs.writeFileSync(
    statePath,
    `${JSON.stringify(
      {
        last_run_at: checkedAt,
        last_run_at_local: record.checkedAtLocal,
        last_cash: dashboard.cash,
        last_cash_number: record.cashNumber,
        last_dashboard_day: dashboard.day,
        last_warehouse_inventory: currentInventory,
        last_warehouse_day: inventoryTable.latestWarehouse.day,
        last_threshold: threshold,
        last_notified_at: notified
          ? checkedAt
          : previousState.last_notified_at || null,
        last_notified_at_local: notified
          ? record.checkedAtLocal
          : previousState.last_notified_at_local || null,
      },
      null,
      2,
    )}\n`,
    "utf8",
  );

  console.log(`Checked at: ${record.checkedAtLocal} (${config.crawl.timezone})`);
  console.log(`Cash: ${dashboard.cash}`);
  console.log(`Dashboard day: ${dashboard.day}`);
  console.log(`Warehouse inventory: ${currentInventory}`);
  console.log(`Warehouse inventory day: ${inventoryTable.latestWarehouse.day}`);
  console.log(`Threshold: ${threshold}`);
  console.log(`Notification sent: ${notified ? "yes" : "no"}`);
  console.log(`State: ${statePath}`);
  console.log(`History: ${historyPath}`);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runWatch(config) {
  const intervalMinutes = Number(config.crawl.interval_minutes || 15);

  if (!Number.isFinite(intervalMinutes) || intervalMinutes <= 0) {
    throw new Error("crawl.interval_minutes must be a positive number");
  }

  while (true) {
    await runOnce(config);
    await sleep(intervalMinutes * 60 * 1000);
  }
}

async function main() {
  const config = readJson(CONFIG_PATH);

  if (!config) {
    throw new Error(`Missing config file: ${CONFIG_PATH}`);
  }

  if (process.argv.includes("--watch")) {
    await runWatch(config);
    return;
  }

  await runOnce(config);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});
