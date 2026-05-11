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

function formatSignedCurrency(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  if (value === 0) {
    return "$0.00";
  }

  return `${value > 0 ? "+" : "-"}$${formatNumber(Math.abs(value), 2)}`;
}

function formatSignedPercent(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  if (value === 0) {
    return "0.00%";
  }

  return `${value > 0 ? "+" : "-"}${Math.abs(value).toFixed(2)}%`;
}

function stripTags(html) {
  return decodeHtml(html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " "));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function extractPlotLines(html) {
  const linePattern =
    /\{label:\s*'([^']+)'\s*,\s*name:\s*'([^']+)'\s*,\s*points:\s*'([^']*)'/g;
  const lines = [];

  for (const match of html.matchAll(linePattern)) {
    const pointTexts = match[3]
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    const points = pointTexts.map(Number);

    if (points.length % 2 !== 0 || points.some((point) => Number.isNaN(point))) {
      throw new Error(`Invalid points data for ${match[1]}`);
    }

    lines.push({
      label: decodeHtml(match[1]),
      name: match[2],
      points,
      pointTexts,
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

function extractDecimalRule(html, key, fallbackPrecision) {
  const blockMatch = html.match(
    new RegExp(`['"]${key}['"]\\s*:\\s*\\{([\\s\\S]*?)\\}`, "i"),
  );
  const block = blockMatch ? blockMatch[1] : "";
  const precisionMatch = block.match(/decimalPrecision\s*:\s*(\d+)/i);
  const groupingSizeMatch = block.match(/groupingSize\s*:\s*(\d+)/i);
  const groupingSeparatorMatch = block.match(/groupingSeparator\s*:\s*'([^']*)'/i);
  const decimalSeparatorMatch = block.match(/decimalSeparator\s*:\s*'([^']*)'/i);

  return {
    decimalPrecision: precisionMatch
      ? Number.parseInt(precisionMatch[1], 10)
      : fallbackPrecision,
    groupingSeparator: groupingSeparatorMatch ? groupingSeparatorMatch[1] : ",",
    groupingSize: groupingSizeMatch ? Number.parseInt(groupingSizeMatch[1], 10) : 3,
    decimalSeparator: decimalSeparatorMatch ? decimalSeparatorMatch[1] : ".",
  };
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

function formatByRule(value, rule) {
  const text = formatNumber(
    value,
    rule.decimalPrecision,
    rule.groupingSeparator,
    rule.groupingSize,
  );

  return rule.decimalSeparator === "." ? text : text.replace(".", rule.decimalSeparator);
}

function csvValue(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function buildPlotRows(lines, xRule, yRule) {
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
    const sourceRow = Array(lines.length + 1).fill("");
    const formattedRow = Array(lines.length + 1).fill("");
    rawRow[0] = minX;
    sourceRow[0] = lines[selected[0]].pointTexts[indexes[selected[0]]];
    formattedRow[0] = formatByRule(minX, xRule);

    for (const lineIndex of selected) {
      const value = lines[lineIndex].points[indexes[lineIndex] + 1];
      rawRow[lineIndex + 1] = value;
      sourceRow[lineIndex + 1] = lines[lineIndex].pointTexts[indexes[lineIndex] + 1];
      formattedRow[lineIndex + 1] = formatByRule(value, yRule);
      indexes[lineIndex] += 2;
    }

    rows.push({ raw: rawRow, source: sourceRow, formatted: formattedRow });
    consumed += selected.length;
  }

  return rows;
}

function parseWarehouseInventoryTable(html) {
  const lines = extractPlotLines(html);
  const header = [extractAxisLabel(html), ...lines.map((line) => line.label)];
  const rows = buildPlotRows(
    lines,
    extractDecimalRule(html, "x-data", 3),
    extractDecimalRule(html, "y-data", 0),
  );
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

function metricKey(sourceId, seriesName) {
  return `${sourceId}:${String(seriesName || "value").trim() || "value"}`;
}

function formatDelta(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  if (value === 0) {
    return "0";
  }

  const abs = Math.abs(value);
  const precision = abs >= 100 ? 2 : abs >= 1 ? 4 : 6;
  return `${value > 0 ? "+" : "-"}${abs.toFixed(precision).replace(/\.?0+$/, "")}`;
}

function formatChangePercent(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  if (value === 0) {
    return "0.00%";
  }

  return `${value > 0 ? "+" : "-"}${Math.abs(value).toFixed(2)}%`;
}

function parsePlotSnapshot(html, source) {
  const lines = extractPlotLines(html);
  const xRule = extractDecimalRule(html, "x-data", 3);
  const yRule = extractDecimalRule(html, "y-data", 0);
  const rows = buildPlotRows(lines, xRule, yRule);
  const series = lines.map((line) => {
    const lastIndex = line.points.length - 2;
    const dayNumber = line.points[lastIndex];
    const valueNumber = line.points[lastIndex + 1];

    return {
      key: metricKey(source.id, line.name || line.label),
      section: source.section,
      plot: source.label,
      series: line.label.trim() || "value",
      day: formatByRule(dayNumber, xRule),
      dayRaw: line.pointTexts[lastIndex],
      value: formatByRule(valueNumber, yRule),
      valueRaw: line.pointTexts[lastIndex + 1],
      valueNumber,
    };
  });

  return {
    id: source.id,
    section: source.section,
    label: source.label,
    axisLabel: extractAxisLabel(html),
    header: [extractAxisLabel(html), ...lines.map((line) => line.label)],
    rows,
    series,
  };
}

function buildOperationalSnapshot(plotSnapshots, previousMetrics = {}) {
  const metrics = {};
  const sections = [];
  const sectionMap = new Map();

  for (const plot of plotSnapshots) {
    if (!sectionMap.has(plot.section)) {
      const section = { name: plot.section, plots: [] };
      sectionMap.set(plot.section, section);
      sections.push(section);
    }

    const enrichedSeries = plot.series.map((series) => {
      const previous = previousMetrics[series.key] || null;
      const previousValueNumber =
        previous && Number.isFinite(previous.valueNumber)
          ? previous.valueNumber
          : null;
      const delta =
        previousValueNumber === null ? null : series.valueNumber - previousValueNumber;
      const changePercent =
        previousValueNumber === null || previousValueNumber === 0
          ? null
          : (delta / Math.abs(previousValueNumber)) * 100;

      const enriched = {
        ...series,
        previousValueRaw: previous ? previous.valueRaw : "",
        previousDayRaw: previous ? previous.dayRaw : "",
        delta,
        deltaText: delta === null ? "" : formatDelta(delta),
        changePercent,
        changePercentText:
          changePercent === null ? "" : formatChangePercent(changePercent),
      };

      metrics[series.key] = {
        section: series.section,
        plot: series.plot,
        series: series.series,
        dayRaw: series.dayRaw,
        valueRaw: series.valueRaw,
        valueNumber: series.valueNumber,
      };

      return enriched;
    });

    sectionMap.get(plot.section).plots.push({
      ...plot,
      series: enrichedSeries,
    });
  }

  return { sections, metrics };
}

function buildOperationalSnapshotCsv(snapshot) {
  const header = [
    "section",
    "plot",
    "series",
    "day_raw",
    "current_value_raw",
    "previous_value_raw",
    "delta",
    "change_percent",
  ];
  const rows = [];

  for (const section of snapshot.sections) {
    for (const plot of section.plots) {
      for (const series of plot.series) {
        rows.push([
          section.name,
          plot.label,
          series.series,
          series.dayRaw,
          series.valueRaw,
          series.previousValueRaw,
          series.deltaText,
          series.changePercentText,
        ]);
      }
    }
  }

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
}

function parseStandingTable(html, targetTeam) {
  const tableMatch = html.match(
    /<table[^>]*id=['"]?standingTable['"]?[^>]*>([\s\S]*?)<\/table>/i,
  );

  if (!tableMatch) {
    throw new Error("Could not find standing table");
  }

  const rowPattern =
    /<tr[^>]*>\s*<td[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*>([\s\S]*?)<\/td>\s*<\/tr>/gi;
  const standings = [];

  for (const match of tableMatch[1].matchAll(rowPattern)) {
    const rankText = stripTags(match[1]);
    const team = stripTags(match[2]);
    const cash = stripTags(match[3]);
    const rank = Number.parseInt(rankText, 10);
    const cashNumber = extractCashNumber(cash);

    if (!Number.isFinite(rank) || !team || cashNumber === null) {
      continue;
    }

    standings.push({
      rank,
      team,
      cash,
      cashNumber,
    });
  }

  if (standings.length === 0) {
    throw new Error("Standing table did not contain team rows");
  }

  const normalizedTarget = targetTeam.toLowerCase();
  const target = standings.find(
    (standing) => standing.team.toLowerCase() === normalizedTarget,
  );

  if (!target) {
    throw new Error(`Could not find target team in standing table: ${targetTeam}`);
  }

  const rows = standings.map((standing) => {
    const gapAmount = target.cashNumber - standing.cashNumber;
    const gapPercent =
      standing.cashNumber === 0 ? null : (gapAmount / standing.cashNumber) * 100;

    return {
      ...standing,
      gapAmount,
      gapAmountText: formatSignedCurrency(gapAmount),
      gapPercent,
      gapPercentText: formatSignedPercent(gapPercent),
    };
  });

  return {
    target,
    rows,
  };
}

function buildWarehouseCsv({ header, rows }) {
  const csvLines = [
    header.map(csvValue).join(","),
    ...rows.map((row) => row.formatted.map(csvValue).join(",")),
  ];

  return `\uFEFF${csvLines.join("\r\n")}\r\n`;
}

function buildStandingGapsCsv({ rows }) {
  const header = [
    "rank",
    "team",
    "cash",
    "cash_number",
    "gap_amount_target_minus_team",
    "gap_percent_vs_team_cash",
  ];
  const csvRows = rows.map((row) => [
    row.rank,
    row.team,
    row.cash,
    row.cashNumber,
    row.gapAmountText,
    row.gapPercentText,
  ]);

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...csvRows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
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
    "target_team",
    "target_rank",
    "target_cash",
    "target_cash_number",
    "threshold",
    "inventory_alert",
    "email_sent",
  ];
  const row = [
    record.checkedAt,
    record.cash,
    record.cashNumber ?? "",
    record.dashboardDay,
    record.warehouseInventory,
    record.warehouseDay,
    record.targetTeam,
    record.targetRank,
    record.targetCash,
    record.targetCashNumber ?? "",
    record.threshold,
    record.inventoryAlert ? "yes" : "no",
    record.emailSent ? "yes" : "no",
  ];
  const headerLine = header.map(csvValue).join(",");
  const line = row.map(csvValue).join(",");
  let prefix = "";

  if (!exists) {
    prefix = `${headerLine}\n`;
  } else {
    const firstLine = fs.readFileSync(filePath, "utf8").split(/\r?\n/, 1)[0];

    if (firstLine !== headerLine) {
      fs.renameSync(filePath, `${filePath}.bak-${Date.now()}`);
      prefix = `${headerLine}\n`;
    }
  }

  fs.appendFileSync(filePath, `${prefix}${line}\n`, "utf8");
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

function isEmailDryRun() {
  return ["1", "true", "yes"].includes(
    optionalEnv("EMAIL_DRY_RUN", "false").toLowerCase(),
  );
}

function isEmailEnabled(config) {
  const envOverride = optionalEnv("EMAIL_ENABLED", "").toLowerCase();

  if (["0", "false", "no"].includes(envOverride)) {
    return false;
  }

  return config.email.enabled;
}

function buildStandingLines(standingReport) {
  return standingReport.rows.map((row) =>
    [
      `rank ${row.rank}`,
      row.team,
      row.cash,
      row.gapAmountText,
      row.gapPercentText,
    ].join(" | "),
  );
}

function buildOperationalLines(snapshot) {
  if (!snapshot) {
    return [];
  }

  const lines = [];

  for (const section of snapshot.sections) {
    lines.push("");
    lines.push(`[${section.name}]`);

    for (const plot of section.plots) {
      for (const series of plot.series) {
        lines.push(
          [
            plot.label,
            series.series,
            `day ${series.dayRaw}`,
            `current ${series.valueRaw}`,
            `1h ${series.deltaText || "n/a"}`,
            `${series.changePercentText || "n/a"}`,
          ].join(" | "),
        );
      }
    }
  }

  return lines;
}

function findOperationalMetric(snapshot, key) {
  if (!snapshot) {
    return null;
  }

  for (const section of snapshot.sections) {
    for (const plot of section.plots) {
      const found = plot.series.find((series) => series.key === key);

      if (found) {
        return found;
      }
    }
  }

  return null;
}

function buildMetricAlertLines(alerts = []) {
  if (alerts.length === 0) {
    return [];
  }

  return [
    "",
    "[Metric Alerts]",
    ...alerts.map((alert) =>
      [
        alert.label,
        `current ${alert.currentRaw}`,
        `${alert.direction} ${alert.threshold}`,
      ].join(" | "),
    ),
  ];
}

function collectMetricAlerts(config, snapshot) {
  if (!snapshot) {
    return [];
  }

  const thresholdConfig = config.monitor.metric_thresholds || {};
  const alerts = [];

  for (const [key, threshold] of Object.entries(thresholdConfig)) {
    if (!threshold || typeof threshold !== "object") {
      continue;
    }

    const metric = findOperationalMetric(snapshot, key);

    if (!metric || !Number.isFinite(metric.valueNumber)) {
      continue;
    }

    const min =
      threshold.min === null || threshold.min === undefined
        ? null
        : Number(threshold.min);
    const max =
      threshold.max === null || threshold.max === undefined
        ? null
        : Number(threshold.max);
    const label =
      threshold.label ||
      `${metric.section} ${metric.plot} ${metric.series}`.replace(/\s+/g, " ");

    if (Number.isFinite(max) && metric.valueNumber >= max) {
      alerts.push({
        key,
        label,
        currentRaw: metric.valueRaw,
        threshold: max,
        direction: ">= max",
      });
    }

    if (Number.isFinite(min) && metric.valueNumber <= min) {
      alerts.push({
        key,
        label,
        currentRaw: metric.valueRaw,
        threshold: min,
        direction: "<= min",
      });
    }
  }

  return alerts;
}

function metricSummaryLines(snapshot, limit = 12) {
  if (!snapshot) {
    return [];
  }

  const lines = [];

  for (const section of snapshot.sections) {
    for (const plot of section.plots) {
      for (const series of plot.series) {
        lines.push(
          [
            `${series.key}`,
            `${section.name}/${plot.label}/${series.series}`,
            `day=${series.dayRaw}`,
            `value=${series.valueRaw}`,
            `delta_1h=${series.deltaText || "n/a"}`,
            `pct_1h=${series.changePercentText || "n/a"}`,
          ].join(" | "),
        );

        if (lines.length >= limit) {
          return lines;
        }
      }
    }
  }

  return lines;
}

function standingSummaryLines(standingReport, limit = 8) {
  return standingReport.rows.slice(0, limit).map((row) =>
    [
      `rank=${row.rank}`,
      `team=${row.team}`,
      `cash=${row.cash}`,
      `gap=${row.gapAmountText}`,
      `gap_pct=${row.gapPercentText}`,
    ].join(" | "),
  );
}

function buildFallbackRecommendations(config, record, standingReport, options = {}) {
  const suggestions = [];
  const snapshot = options.operationalSnapshot;
  const lostDemand = findOperationalMetric(snapshot, "hq_lost_demand:Calopeia");
  const demand = findOperationalMetric(snapshot, "hq_demand:Calopeia");
  const wip = findOperationalMetric(snapshot, "factory_wip:Calopeia");
  const shipments = findOperationalMetric(snapshot, "warehouse_shipments:Calopeia");

  if (record.inventoryAlert) {
    suggestions.push(
      `倉庫庫存 ${record.warehouseInventory} 已達 ${record.threshold} 預警值；若需求支撐，優先降低入庫或加快出貨。`,
    );
  } else if (record.warehouseInventory <= 0 && lostDemand?.valueNumber > 0) {
    suggestions.push(
      `倉庫庫存為 ${record.warehouseInventory} 且 lost demand 為 ${lostDemand.valueRaw}；優先恢復補貨與可出貨量。`,
    );
  } else {
    suggestions.push(
      `倉庫庫存 ${record.warehouseInventory} 低於 ${record.threshold} 預警值；維持生產與出貨貼近需求。`,
    );
  }

  if (lostDemand?.valueNumber > 0) {
    suggestions.push(
      `Lost demand 為 ${lostDemand.valueRaw}；檢查瓶頸在倉庫庫存、出貨時點或工廠產出。`,
    );
  }

  if (wip?.valueNumber > 0 && shipments?.valueNumber === 0) {
    suggestions.push(
      `Factory WIP 為 ${wip.valueRaw} 且 shipments 為 ${shipments.valueRaw}；增加 WIP 前先檢查出貨流程。`,
    );
  }

  if (standingReport.target.rank === 1) {
    suggestions.push(
      `現金排名第 1、金額 ${record.targetCash}；保護領先時避免過量庫存成本與需求流失。`,
    );
  } else {
    suggestions.push(
      `現金排名第 ${standingReport.target.rank}；採取高成本追趕動作前先比對排名差距表。`,
    );
  }

  return suggestions.slice(0, Number(config.ai?.max_suggestions || 4));
}

function extractRecommendationItems(text, limit) {
  const cleaned = String(text || "")
    .replace(/```(?:json)?/gi, "")
    .replace(/```/g, "")
    .trim();

  try {
    const parsed = JSON.parse(cleaned);
    const items = Array.isArray(parsed) ? parsed : parsed.recommendations;

    if (Array.isArray(items)) {
      return items
        .map((item) => String(item).trim())
        .filter(Boolean)
        .slice(0, limit);
    }
  } catch {
    // Fall through to line parsing when the model returns plain text.
  }

  return cleaned
    .split(/\r?\n/)
    .map((line) => line.replace(/^\s*[-*•\d.)]+\s*/, "").trim())
    .filter(Boolean)
    .slice(0, limit);
}

function geminiApiKey(config) {
  const configuredEnv = config.ai?.api_key_env || "GEMINI_API_KEY";

  return (
    optionalEnv("GOOGLE_API_KEY") ||
    optionalEnv(configuredEnv) ||
    optionalEnv("GEMINI_API_KEY")
  );
}

function buildRecommendationPrompt(config, record, standingReport, options = {}) {
  const metricAlerts = options.metricAlerts || [];
  const lines = [
    "You are advising a team in a supply chain simulation.",
    "Use only the data below. Do not invent missing values.",
    "Write Traditional Chinese. Return 2 to 4 concise action recommendations.",
    "Each recommendation should be one short sentence, useful for operations decisions, and avoid formulas.",
    "Return JSON only in this shape: {\"recommendations\":[\"...\"]}",
    "",
    `Target team: ${record.targetTeam}`,
    `Checked at: ${record.checkedAtLocal} ${config.crawl.timezone}`,
    `Cash: ${record.targetCash}`,
    `Rank: ${record.targetRank}`,
    `Dashboard day: ${record.dashboardDay}`,
    `Warehouse inventory: ${record.warehouseInventory}`,
    `Warehouse inventory day: ${record.warehouseDay}`,
    `Warehouse inventory alert threshold: ${record.threshold}`,
    `Warehouse inventory alert: ${record.inventoryAlert ? "yes" : "no"}`,
    "",
    "Standing rows:",
    ...standingSummaryLines(standingReport),
  ];

  if (options.operationalSnapshot) {
    lines.push("", "Operational metrics:", ...metricSummaryLines(options.operationalSnapshot));
  }

  if (metricAlerts.length > 0) {
    lines.push(
      "",
      "Configured metric alerts:",
      ...metricAlerts.map(
        (alert) =>
          `${alert.label}: current=${alert.currentRaw}, rule=${alert.direction} ${alert.threshold}`,
      ),
    );
  }

  return lines.join("\n");
}

async function buildRecommendations(config, record, standingReport, options = {}) {
  const limit = Number(config.ai?.max_suggestions || 4);
  const fallback = buildFallbackRecommendations(config, record, standingReport, options);

  if (!config.ai?.enabled) {
    return {
      source: "Local rules",
      items: fallback,
    };
  }

  const apiKey = geminiApiKey(config);

  if (!apiKey) {
    console.warn("Gemini recommendations skipped: GEMINI_API_KEY is not set");
    return {
      source: "Local rules",
      items: fallback,
    };
  }

  try {
    const { GoogleGenAI } = await import("@google/genai");
    const model = config.ai.model || "gemini-2.5-flash";
    const ai = new GoogleGenAI({ apiKey });
    const response = await ai.models.generateContent({
      model,
      contents: buildRecommendationPrompt(config, record, standingReport, options),
    });
    const responseText =
      typeof response.text === "function" ? response.text() : response.text;
    const items = extractRecommendationItems(responseText, limit);

    if (items.length === 0) {
      throw new Error("Gemini returned no recommendation items");
    }

    return {
      source: `Gemini ${model}`,
      items,
    };
  } catch (error) {
    console.warn(`Gemini recommendations failed: ${error.message}`);
    return {
      source: "Local rules",
      items: fallback,
    };
  }
}

function buildEmailSubject(config, record, options = {}) {
  const testPrefix = options.test ? "[TEST] " : "";
  const alertPrefix = record.inventoryAlert ? "ALERT " : "";

  if (options.kind === "warning") {
    const minutes = options.warningMinutes || config.monitor.warning_minutes || 5;
    return `${testPrefix}${config.email.subject_prefix}: ${minutes}-minute warning - ${record.targetTeam} ${record.targetCash} day ${record.dashboardDay}`;
  }

  return `${testPrefix}${alertPrefix}${config.email.subject_prefix}: hourly report - ${record.targetTeam} ${record.targetCash} day ${record.dashboardDay}`;
}

function buildReportText(config, record, standingReport, options = {}) {
  const reportName =
    options.kind === "warning"
      ? `${options.warningMinutes || config.monitor.warning_minutes || 5}-minute warning`
      : "hourly report";

  return [
    `${options.test ? "[TEST] " : ""}${config.email.subject_prefix} ${reportName}`,
    "",
    `Checked at: ${record.checkedAtLocal} (${config.crawl.timezone})`,
    `Target team: ${record.targetTeam}`,
    `Target rank: ${record.targetRank}`,
    `Target cash: ${record.targetCash}`,
    `Dashboard day: ${record.dashboardDay}`,
    `Warehouse inventory: ${record.warehouseInventory}`,
    `Warehouse inventory day: ${record.warehouseDay}`,
    `Warehouse threshold: ${record.threshold}`,
    `Inventory alert: ${record.inventoryAlert ? "YES" : "no"}`,
    "",
    "Recommendations",
    ...((options.recommendations?.items || []).map((item) => `- ${item}`)),
    `Source: ${options.recommendations?.source || "n/a"}`,
    ...buildMetricAlertLines(options.metricAlerts),
    ...buildOperationalLines(options.operationalSnapshot),
    "",
    "Rank | Team | Cash | Gap amount | Gap percent",
    ...buildStandingLines(standingReport),
  ].join("\n");
}

function buildCard(label, value, accent = "#2563eb") {
  return [
    '<td style="padding:8px;width:25%;">',
    `<div style="border:1px solid #d8dee9;border-left:4px solid ${accent};border-radius:8px;padding:12px;background:#ffffff;">`,
    `<div style="font-size:12px;color:#64748b;text-transform:uppercase;letter-spacing:.04em;">${escapeHtml(label)}</div>`,
    `<div style="font-size:20px;font-weight:700;color:#0f172a;margin-top:4px;">${escapeHtml(value)}</div>`,
    "</div>",
    "</td>",
  ].join("");
}

function buildStandingRowsHtml(standingReport) {
  return standingReport.rows
    .map((row) => {
      const target = row.team === standingReport.target.team;
      const gapColor =
        row.gapAmount > 0 ? "#047857" : row.gapAmount < 0 ? "#b91c1c" : "#334155";
      const background = target ? "#eef2ff" : "#ffffff";

      return [
        `<tr style="background:${background};">`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(row.rank)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:${target ? "700" : "500"};">${escapeHtml(row.team)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(row.cash)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;color:${gapColor};font-weight:700;">${escapeHtml(row.gapAmountText)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;color:${gapColor};font-weight:700;">${escapeHtml(row.gapPercentText)}</td>`,
        "</tr>",
      ].join("");
    })
    .join("");
}

function changeColor(value) {
  if (!Number.isFinite(value)) {
    return "#64748b";
  }

  if (value > 0) {
    return "#047857";
  }

  if (value < 0) {
    return "#b91c1c";
  }

  return "#334155";
}

function buildOperationalRowsHtml(snapshot) {
  if (!snapshot) {
    return "";
  }

  return snapshot.sections
    .flatMap((section) =>
      section.plots.flatMap((plot) =>
        plot.series.map((series) => {
          const color = changeColor(series.delta);

          return [
            "<tr>",
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:700;">${escapeHtml(section.name)}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(plot.label)}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(series.series)}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(series.dayRaw)}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;">${escapeHtml(series.valueRaw)}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;color:${color};font-weight:700;">${escapeHtml(series.deltaText || "n/a")}</td>`,
            `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;color:${color};font-weight:700;">${escapeHtml(series.changePercentText || "n/a")}</td>`,
            "</tr>",
          ].join("");
        }),
      ),
    )
    .join("");
}

function buildOperationalSnapshotHtml(snapshot) {
  if (!snapshot) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Operations Snapshot</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
    '<thead><tr style="background:#e2e8f0;color:#334155;">',
    '<th style="padding:9px;text-align:left;">Area</th>',
    '<th style="padding:9px;text-align:left;">Metric</th>',
    '<th style="padding:9px;text-align:left;">Series</th>',
    '<th style="padding:9px;text-align:right;">Day</th>',
    '<th style="padding:9px;text-align:right;">Current</th>',
    '<th style="padding:9px;text-align:right;">1h Change</th>',
    '<th style="padding:9px;text-align:right;">1h %</th>',
    "</tr></thead>",
    `<tbody>${buildOperationalRowsHtml(snapshot)}</tbody>`,
    "</table>",
    "</div>",
  ].join("");
}

function buildRecommendationsHtml(recommendations) {
  const items = recommendations?.items || [];

  if (items.length === 0) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Action Notes</div>',
    '<div style="border:1px solid #bfdbfe;border-radius:8px;background:#eff6ff;padding:12px 14px;">',
    '<ol style="margin:0;padding-left:20px;color:#0f172a;font-size:14px;line-height:1.55;">',
    ...items.map((item) => `<li style="margin:4px 0;">${escapeHtml(item)}</li>`),
    "</ol>",
    `<div style="font-size:11px;color:#64748b;margin-top:8px;">Source: ${escapeHtml(recommendations.source || "n/a")}</div>`,
    "</div>",
    "</div>",
  ].join("");
}

function buildMetricAlertsHtml(alerts = []) {
  if (alerts.length === 0) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Configured Alerts</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #fecaca;border-radius:8px;overflow:hidden;font-size:13px;">',
    '<thead><tr style="background:#fee2e2;color:#7f1d1d;">',
    '<th style="padding:9px;text-align:left;">Metric</th>',
    '<th style="padding:9px;text-align:right;">Current</th>',
    '<th style="padding:9px;text-align:right;">Rule</th>',
    "</tr></thead>",
    "<tbody>",
    ...alerts.map((alert) =>
      [
        "<tr>",
        `<td style="padding:8px;border-bottom:1px solid #fecaca;font-weight:700;">${escapeHtml(alert.label)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #fecaca;text-align:right;">${escapeHtml(alert.currentRaw)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #fecaca;text-align:right;color:#b91c1c;font-weight:700;">${escapeHtml(`${alert.direction} ${alert.threshold}`)}</td>`,
        "</tr>",
      ].join(""),
    ),
    "</tbody>",
    "</table>",
    "</div>",
  ].join("");
}

function buildEmailFooterHtml(config) {
  const footer = config.email.footer || {};
  const text = footer.text || "Developed by Yung-Sian Fang";
  const url = footer.url || "https://sammyfang.tw";

  return [
    '<div style="padding:14px 22px;background:#0f172a;color:#cbd5e1;font-size:12px;border-top:1px solid #1e293b;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">',
    "<tr>",
    `<td style="vertical-align:middle;">${escapeHtml(text)}</td>`,
    `<td style="vertical-align:middle;text-align:right;"><a href="${escapeHtml(url)}" style="color:#93c5fd;text-decoration:none;font-weight:700;">${escapeHtml(url.replace(/^https?:\/\//, ""))}</a></td>`,
    "</tr>",
    "</table>",
    "</div>",
  ].join("");
}

function buildReportHtml(config, record, standingReport, options = {}) {
  const isWarning = options.kind === "warning";
  const minutes = options.warningMinutes || config.monitor.warning_minutes || 5;
  const title = isWarning
    ? `${minutes}-minute warning`
    : "Hourly monitoring report";
  const eyebrow = options.test ? "Test email" : "Supply Chain Watchdog";
  const alertText = record.inventoryAlert
    ? `Warehouse inventory is at or above ${record.threshold}.`
    : `Warehouse inventory is below ${record.threshold}.`;
  const bannerColor = isWarning || record.inventoryAlert ? "#b91c1c" : "#1d4ed8";
  const bannerBg = isWarning || record.inventoryAlert ? "#fef2f2" : "#eff6ff";

  return [
    '<!doctype html>',
    '<html><body style="margin:0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;color:#0f172a;">',
    '<div style="max-width:820px;margin:0 auto;padding:24px;">',
    '<div style="background:#ffffff;border:1px solid #d8dee9;border-radius:12px;overflow:hidden;">',
    `<div style="background:${bannerColor};color:#ffffff;padding:18px 22px;">`,
    `<div style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;opacity:.9;">${escapeHtml(eyebrow)}</div>`,
    `<div style="font-size:24px;font-weight:700;margin-top:4px;">${escapeHtml(title)}</div>`,
    `<div style="font-size:13px;margin-top:6px;opacity:.95;">${escapeHtml(record.checkedAtLocal)} (${escapeHtml(config.crawl.timezone)})</div>`,
    "</div>",
    `<div style="padding:16px 22px;background:${bannerBg};border-bottom:1px solid #d8dee9;color:${bannerColor};font-weight:700;">${escapeHtml(alertText)}</div>`,
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="padding:14px;">',
    "<tr>",
    buildCard("Target team", record.targetTeam),
    buildCard("Rank", record.targetRank),
    buildCard("Cash", record.targetCash, "#047857"),
    buildCard("Day", record.dashboardDay, "#7c3aed"),
    "</tr><tr>",
    buildCard("Warehouse inventory", record.warehouseInventory, record.inventoryAlert ? "#b91c1c" : "#0f766e"),
    buildCard("Inventory day", record.warehouseDay),
    buildCard("Threshold", record.threshold, "#ea580c"),
    buildCard("Report type", isWarning ? `${minutes}-minute warning` : "hourly"),
    "</tr>",
    "</table>",
    buildRecommendationsHtml(options.recommendations),
    buildMetricAlertsHtml(options.metricAlerts),
    !isWarning ? buildOperationalSnapshotHtml(options.operationalSnapshot) : "",
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Team Standing</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:14px;">',
    '<thead><tr style="background:#e2e8f0;color:#334155;">',
    '<th style="padding:9px;text-align:right;">Rank</th>',
    '<th style="padding:9px;text-align:left;">Team</th>',
    '<th style="padding:9px;text-align:right;">Cash</th>',
    '<th style="padding:9px;text-align:right;">Gap amount</th>',
    '<th style="padding:9px;text-align:right;">Gap percent</th>',
    "</tr></thead>",
    `<tbody>${buildStandingRowsHtml(standingReport)}</tbody>`,
    "</table>",
    "</div>",
    buildEmailFooterHtml(config),
    "</div></div></body></html>",
  ].join("");
}

async function sendReportEmail(config, record, standingReport, options = {}) {
  if (!isEmailEnabled(config)) {
    console.log("Email report skipped because email is disabled");
    return false;
  }

  const recipients = config.email.recipients || [];

  if (recipients.length === 0) {
    throw new Error("Email report is enabled, but email.recipients is empty");
  }

  const metricAlerts =
    options.metricAlerts ?? collectMetricAlerts(config, options.operationalSnapshot);
  const recommendations =
    options.recommendations ??
    (await buildRecommendations(config, record, standingReport, {
      ...options,
      metricAlerts,
    }));
  const renderOptions = {
    ...options,
    metricAlerts,
    recommendations,
  };
  const subject = buildEmailSubject(config, record, renderOptions);
  const text = buildReportText(config, record, standingReport, renderOptions);
  const html = buildReportHtml(config, record, standingReport, renderOptions);

  if (isEmailDryRun()) {
    console.log("EMAIL_DRY_RUN=1, email report not sent.");
    console.log(`Dry-run recipients: ${recipients.join(", ")}`);
    console.log(`Dry-run subject: ${subject}`);
    console.log(`Recommendations source: ${recommendations.source}`);
    for (const item of recommendations.items) {
      console.log(`- ${item}`);
    }
    return false;
  }

  const transporter = createTransport();

  await transporter.sendMail({
    from: requiredEnv("SMTP_FROM"),
    to: recipients.join(","),
    subject,
    text,
    html,
  });

  return true;
}

async function crawl(config, options = {}) {
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
  const plotSnapshots = [];

  if (options.includePlots !== false) {
    for (const source of config.crawl.plot_sources || []) {
      let plotHtml = inventoryHtml;

      if (source.url !== config.crawl.warehouse_inventory_url) {
        const plotResponse = await request(source.url, { method: "GET" }, cookieJar);
        plotHtml = await plotResponse.text();

        if (!plotResponse.ok) {
          throw new Error(`${source.label} plot failed with HTTP ${plotResponse.status}`);
        }
      }

      plotSnapshots.push(parsePlotSnapshot(plotHtml, source));
    }
  }

  const standingResponse = await request(
    config.crawl.standing_url,
    { method: "POST" },
    cookieJar,
  );
  const standingHtml = await standingResponse.text();

  if (!standingResponse.ok) {
    throw new Error(`Standing table failed with HTTP ${standingResponse.status}`);
  }

  const standingReport = parseStandingTable(standingHtml, config.monitor.target_team);

  return {
    dashboard,
    inventoryTable,
    plotSnapshots,
    standingReport,
  };
}

function createRecord(config, dashboard, inventoryTable, standingReport, options = {}) {
  const checkedAt = options.checkedAt || new Date().toISOString();
  const threshold = Number(config.monitor.warehouse_inventory_threshold);
  const currentInventory = inventoryTable.latestWarehouse.inventory;
  const inventoryAlert =
    options.inventoryAlertOverride ?? currentInventory >= threshold;

  return {
    checkedAt,
    checkedAtLocal: localizedTime(checkedAt, config.crawl.timezone),
    cash: dashboard.cash,
    cashNumber: extractCashNumber(dashboard.cash),
    dashboardDay: dashboard.day,
    warehouseInventory: currentInventory,
    warehouseDay: inventoryTable.latestWarehouse.day,
    targetTeam: standingReport.target.team,
    targetRank: standingReport.target.rank,
    targetCash: standingReport.target.cash,
    targetCashNumber: standingReport.target.cashNumber,
    threshold,
    inventoryAlert,
  };
}

async function runOnce(config) {
  const statePath = path.resolve(process.cwd(), config.output.latest_json);
  const historyPath = path.resolve(process.cwd(), config.output.history_csv);
  const inventoryCsvPath = path.resolve(
    process.cwd(),
    config.output.warehouse_inventory_csv,
  );
  const standingCsvPath = path.resolve(
    process.cwd(),
    config.output.standing_gaps_csv,
  );
  const operationalCsvPath = path.resolve(
    process.cwd(),
    config.output.operational_snapshot_csv,
  );
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));

  const previousState = readJson(statePath, {});
  const { dashboard, inventoryTable, plotSnapshots, standingReport } =
    await crawl(config);
  const operationalSnapshot = buildOperationalSnapshot(
    plotSnapshots,
    previousState.last_operational_metrics || {},
  );
  const checkedAt = new Date().toISOString();
  const record = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });
  let emailSent = false;

  if (config.monitor.send_report_every_run) {
    emailSent = await sendReportEmail(config, record, standingReport, {
      operationalSnapshot,
    });
  }

  fs.writeFileSync(inventoryCsvPath, buildWarehouseCsv(inventoryTable), "utf8");
  fs.writeFileSync(standingCsvPath, buildStandingGapsCsv(standingReport), "utf8");
  fs.writeFileSync(
    operationalCsvPath,
    buildOperationalSnapshotCsv(operationalSnapshot),
    "utf8",
  );
  appendHistory(historyPath, { ...record, emailSent });
  fs.writeFileSync(
    statePath,
    `${JSON.stringify(
      {
        last_run_at: checkedAt,
        last_run_at_local: record.checkedAtLocal,
        last_cash: dashboard.cash,
        last_cash_number: record.cashNumber,
        last_dashboard_day: dashboard.day,
        last_warehouse_inventory: record.warehouseInventory,
        last_warehouse_day: inventoryTable.latestWarehouse.day,
        last_target_team: standingReport.target.team,
        last_target_rank: standingReport.target.rank,
        last_target_cash: standingReport.target.cash,
        last_target_cash_number: standingReport.target.cashNumber,
        last_threshold: record.threshold,
        last_inventory_alert: record.inventoryAlert,
        last_operational_metrics: operationalSnapshot.metrics,
        last_email_sent_at: emailSent
          ? checkedAt
          : previousState.last_email_sent_at || null,
        last_email_sent_at_local: emailSent
          ? record.checkedAtLocal
          : previousState.last_email_sent_at_local || null,
      },
      null,
      2,
    )}\n`,
    "utf8",
  );

  console.log(`Checked at: ${record.checkedAtLocal} (${config.crawl.timezone})`);
  console.log(`Cash: ${dashboard.cash}`);
  console.log(`Dashboard day: ${dashboard.day}`);
  console.log(`Target team: ${standingReport.target.team}`);
  console.log(`Target rank: ${standingReport.target.rank}`);
  console.log(`Target cash: ${standingReport.target.cash}`);
  console.log(`Warehouse inventory: ${record.warehouseInventory}`);
  console.log(`Warehouse inventory day: ${inventoryTable.latestWarehouse.day}`);
  console.log(`Threshold: ${record.threshold}`);
  console.log(`Inventory alert: ${record.inventoryAlert ? "yes" : "no"}`);
  console.log(`Email sent: ${emailSent ? "yes" : "no"}`);
  console.log(`State: ${statePath}`);
  console.log(`History: ${historyPath}`);
  console.log(`Standing gaps: ${standingCsvPath}`);
  console.log(`Operational snapshot: ${operationalCsvPath}`);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runWatch(config) {
  const intervalMinutes = Number(config.crawl.interval_minutes || 60);

  if (!Number.isFinite(intervalMinutes) || intervalMinutes <= 0) {
    throw new Error("crawl.interval_minutes must be a positive number");
  }

  while (true) {
    await runOnce(config);
    await sleep(intervalMinutes * 60 * 1000);
  }
}

async function sendTestEmails(config) {
  const previousState = readJson(
    path.resolve(process.cwd(), config.output.latest_json),
    {},
  );
  const { dashboard, inventoryTable, plotSnapshots, standingReport } =
    await crawl(config);
  const operationalSnapshot = buildOperationalSnapshot(
    plotSnapshots,
    previousState.last_operational_metrics || {},
  );
  const checkedAt = new Date().toISOString();
  const baseRecord = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });

  const hourlySent = await sendReportEmail(config, baseRecord, standingReport, {
    kind: "hourly",
    test: true,
    operationalSnapshot,
  });

  console.log("Test email summary:");
  console.log(`Hourly report email sent: ${hourlySent ? "yes" : "no"}`);
  console.log(`Target team: ${baseRecord.targetTeam}`);
  console.log(`Target cash: ${baseRecord.targetCash}`);
  console.log(`Dashboard day: ${baseRecord.dashboardDay}`);
  console.log(`Warehouse inventory: ${baseRecord.warehouseInventory}`);
  console.log(`Operational metrics: ${Object.keys(operationalSnapshot.metrics).length}`);
}

async function sendWarningEmail(config) {
  const { dashboard, inventoryTable, standingReport } = await crawl(config, {
    includePlots: false,
  });
  const warningMinutes = Number(config.monitor.warning_minutes || 15);
  const record = createRecord(config, dashboard, inventoryTable, standingReport);
  let emailSent = false;

  if (record.inventoryAlert) {
    emailSent = await sendReportEmail(config, record, standingReport, {
      kind: "warning",
      warningMinutes,
    });
  }

  console.log(`${warningMinutes}-minute warning email summary:`);
  console.log(`Email sent: ${emailSent ? "yes" : "no"}`);
  console.log(`Target team: ${record.targetTeam}`);
  console.log(`Target cash: ${record.targetCash}`);
  console.log(`Dashboard day: ${record.dashboardDay}`);
  console.log(`Warehouse inventory: ${record.warehouseInventory}`);
  console.log(`Inventory alert: ${record.inventoryAlert ? "yes" : "no"}`);
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

  if (process.argv.includes("--test-email")) {
    await sendTestEmails(config);
    return;
  }

  if (process.argv.includes("--warning-email")) {
    await sendWarningEmail(config);
    return;
  }

  await runOnce(config);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});
