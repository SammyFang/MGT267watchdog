const fs = require("node:fs");
const path = require("node:path");

require("dotenv").config();

const nodemailer = require("nodemailer");
const ExcelJS = require("exceljs");

const CONFIG_PATH = path.resolve(process.cwd(), "monitor_config.json");
const DEFAULT_GAME_RULES = {
  end_day: 1460,
  product_price_per_drum: 1450,
  customer_fulfillment_cost_per_drum: 150,
  holding_cost_per_drum_per_year: 100,
  cash_interest_rate_annual_percent: 10,
  order_response_hours: 24,
  capacity_expansion_cost_per_drum_per_day: 50000,
  capacity_expansion_lead_days: 90,
  capacity_can_be_retired: false,
  production_fixed_cost_per_batch: 1500,
  production_variable_cost_per_drum: 1000,
  truck_capacity_drums: 200,
  truck_cost: 15000,
  truck_lead_days: 7,
  mail_cost_per_drum: 150,
  mail_lead_days: 1,
  priority_level_affects_supply_chain: false,
};

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

function requestLabel(url) {
  try {
    const parsed = new URL(url);
    return `${parsed.origin}${parsed.pathname}`;
  } catch {
    return String(url);
  }
}

function retryDelayMs(attempt) {
  return Math.min(1500 * 2 ** (attempt - 1), 10000);
}

function isRetryableStatus(status) {
  return [408, 429, 500, 502, 503, 504].includes(status);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function request(url, options, cookieJar) {
  const maxAttempts = Math.max(
    1,
    Number.parseInt(optionalEnv("SC_FETCH_ATTEMPTS", "4"), 10) || 4,
  );
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const headers = new Headers(options.headers || {});
    const cookies = cookieHeader(cookieJar);

    if (cookies) {
      headers.set("cookie", cookies);
    }

    try {
      const response = await fetch(url, {
        ...options,
        headers,
        redirect: "manual",
      });

      collectCookies(response.headers, cookieJar);

      if (!isRetryableStatus(response.status) || attempt === maxAttempts) {
        return response;
      }

      console.warn(
        `Retrying ${requestLabel(url)} after HTTP ${response.status} (${attempt}/${maxAttempts})`,
      );
    } catch (error) {
      lastError = error;

      if (attempt === maxAttempts) {
        throw error;
      }

      console.warn(
        `Retrying ${requestLabel(url)} after fetch error: ${error.message} (${attempt}/${maxAttempts})`,
      );
    }

    await delay(retryDelayMs(attempt));
  }

  throw lastError || new Error(`Request failed: ${requestLabel(url)}`);
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

function htmlAttribute(html, name) {
  const match = String(html || "").match(
    new RegExp(`${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, "i"),
  );

  if (!match) {
    return "";
  }

  return decodeHtml(match[1] || match[2] || match[3] || "");
}

function compactText(value) {
  return decodeHtml(
    String(value || "")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<br\s*\/?>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " "),
  );
}

function selectedOption(selectHtml) {
  const options = [...String(selectHtml || "").matchAll(/<option\b([^>]*)>([\s\S]*?)<\/option>/gi)];

  if (options.length === 0) {
    return { value: "", text: "" };
  }

  const selected =
    options.find((option) => /\bselected\b/i.test(option[1])) || options[0];

  return {
    value: htmlAttribute(selected[1], "value"),
    text: compactText(selected[2]),
  };
}

function selectOptions(selectHtml) {
  return [...String(selectHtml || "").matchAll(/<option\b([^>]*)>([\s\S]*?)<\/option>/gi)].map(
    (option) => ({
      value: htmlAttribute(option[1], "value"),
      text: compactText(option[2]),
      selected: /\bselected\b/i.test(option[1]),
    }),
  );
}

function cellTextWithControls(html) {
  let normalized = String(html || "")
    .replace(/<select\b[\s\S]*?<\/select>/gi, (selectHtml) => {
      const selected = selectedOption(selectHtml);
      return selected.text || selected.value;
    })
    .replace(/<input\b[^>]*>/gi, (inputHtml) => {
      const type = htmlAttribute(inputHtml, "type").toLowerCase();
      const value = htmlAttribute(inputHtml, "value");

      if (type === "checkbox" || type === "radio") {
        const checked = /\bchecked\b/i.test(inputHtml);
        return `${checked ? "yes" : "no"}${value ? ` (${value})` : ""}`;
      }

      return value;
    })
    .replace(/<textarea\b[^>]*>([\s\S]*?)<\/textarea>/gi, (_match, text) =>
      compactText(text),
    );

  return compactText(normalized);
}

function extractFormControls(html) {
  const controls = [];
  const forms = [...String(html || "").matchAll(/<form\b([^>]*)>([\s\S]*?)<\/form>/gi)];

  for (const [formIndex, formMatch] of forms.entries()) {
    const formAttrs = formMatch[1];
    const formHtml = formMatch[2];
    const formAction = htmlAttribute(formAttrs, "action");
    const formMethod = htmlAttribute(formAttrs, "method") || "GET";

    for (const selectMatch of formHtml.matchAll(/<select\b([^>]*)>([\s\S]*?)<\/select>/gi)) {
      const selected = selectedOption(selectMatch[0]);
      controls.push({
        formIndex: formIndex + 1,
        formAction,
        formMethod,
        tag: "select",
        name: htmlAttribute(selectMatch[1], "name"),
        type: "select",
        value: selected.value,
        text: selected.text,
        options: selectOptions(selectMatch[0]),
        checked: "",
      });
    }

    for (const inputMatch of formHtml.matchAll(/<input\b([^>]*)>/gi)) {
      const attrs = inputMatch[1];
      const type = htmlAttribute(attrs, "type") || "text";
      controls.push({
        formIndex: formIndex + 1,
        formAction,
        formMethod,
        tag: "input",
        name: htmlAttribute(attrs, "name"),
        type,
        value: htmlAttribute(attrs, "value"),
        text: "",
        checked:
          type.toLowerCase() === "checkbox" || type.toLowerCase() === "radio"
            ? /\bchecked\b/i.test(attrs)
              ? "yes"
              : "no"
            : "",
      });
    }
  }

  return controls;
}

function inferPolicyTableLabel(rows) {
  const text = rows.flat().join(" ").toLowerCase();

  if (text.includes("warehouse location") && text.includes("shipping method")) {
    return "Production and Shipping";
  }

  if (text.includes("factory location") && text.includes("shipping method")) {
    return "Inbound Shipments from Factories";
  }

  if (text.includes("destination") && text.includes("fulfillment cost")) {
    return "Outbound Shipments to Customers";
  }

  return "Table";
}

function extractHtmlTables(html) {
  const tables = [];

  for (const [index, match] of [
    ...String(html || "").matchAll(/<table\b[^>]*>([\s\S]*?)<\/table>/gi),
  ].entries()) {
    const rows = [];

    for (const rowMatch of match[1].matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
      const cells = [
        ...rowMatch[1].matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi),
      ].map((cellMatch) => cellTextWithControls(cellMatch[1]));

      if (cells.some((cell) => cell !== "")) {
        rows.push(cells);
      }
    }

    if (rows.length > 0) {
      tables.push({
        index: index + 1,
        label: inferPolicyTableLabel(rows),
        rows,
      });
    }
  }

  return tables;
}

function extractPolicyFacts(html) {
  const text = compactText(html);
  const patterns = [
    /Revenue per drum is \$[\d,.]+/i,
    /Each order must be filled within [^.]+?order is lost/i,
    /Your fulfillment policy is currently set to [^.]+?\./i,
    /Factory is operational with a current capacity of [\d.]+/i,
    /No additional capacity is scheduled/i,
    /Cost to produce a batch is \$[\d,.]+ \+ \(batch size\) x \(\$[\d,.]+\)/i,
    /Warehouse is operational/i,
    /Warehouse is under construction and will become operational in [^.]+?\./i,
  ];
  const facts = [];

  for (const pattern of patterns) {
    const match = text.match(pattern);

    if (match) {
      facts.push(match[0].replace(/\s+/g, " ").trim());
    }
  }

  return [...new Set(facts)];
}

function parsePolicyPageSnapshot(html, page) {
  return {
    id: page.id,
    section: page.section,
    region: page.region,
    label: page.label,
    url: page.url,
    facts: extractPolicyFacts(html),
    forms: extractFormControls(html),
    tables: extractHtmlTables(html),
  };
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

function plotSeriesLabel(line) {
  return (
    String(line.label || "").trim() ||
    String(line.name || "").trim() ||
    "value"
  );
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

function safeWorksheetName(name, usedNames) {
  const cleaned = String(name || "Sheet")
    .replace(/[:\\/?*[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 31);
  const base = cleaned || "Sheet";
  let candidate = base;
  let index = 2;

  while (usedNames.has(candidate.toLowerCase())) {
    const suffix = ` ${index}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }

  usedNames.add(candidate.toLowerCase());
  return candidate;
}

function safeAttachmentFilename(value) {
  return String(value || "supply-chain-data")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
}

function dataCell(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value);
}

function columnLetter(index) {
  let dividend = index;
  let column = "";

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    column = String.fromCharCode(65 + modulo) + column;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return column;
}

function styleWorksheet(worksheet) {
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1D4ED8" },
  };
  headerRow.alignment = { vertical: "middle" };

  for (const column of worksheet.columns) {
    let maxLength = 10;

    column.eachCell({ includeEmpty: true }, (cell) => {
      const length = String(cell.value ?? "").length;
      maxLength = Math.max(maxLength, length);
    });

    column.width = Math.min(Math.max(maxLength + 2, 12), 28);
  }
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
  const header = [extractAxisLabel(html), ...lines.map(plotSeriesLabel)];
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
      series: plotSeriesLabel(line),
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
    header: [extractAxisLabel(html), ...lines.map(plotSeriesLabel)],
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

function sortedUniqueNumbers(values) {
  return [...new Set(values.filter(Number.isFinite).map((value) => Number(value)))]
    .sort((a, b) => a - b);
}

function roundMetric(value, digits = 4) {
  if (!Number.isFinite(value)) {
    return null;
  }

  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function percentMetric(value) {
  return Number.isFinite(value) ? `${value.toFixed(2)}%` : "";
}

function ratioMetric(value) {
  if (!Number.isFinite(value)) {
    return "";
  }

  return value.toFixed(3).replace(/\.?0+$/, "");
}

function backtestDayLabel(value, fallback = "") {
  if (!Number.isFinite(value)) {
    return fallback || "";
  }

  const rounded = Math.round(value);

  if (Math.abs(value - rounded) < 0.001) {
    return formatNumber(rounded, 0);
  }

  return value.toFixed(3).replace(/\.?0+$/, "");
}

function seriesPoints(plotSnapshots, sourceId, seriesName) {
  const plot = (plotSnapshots || []).find((item) => item.id === sourceId);

  if (!plot) {
    return [];
  }

  const normalized = String(seriesName || "").trim().toLowerCase();
  const seriesIndex = plot.header
    .slice(1)
    .findIndex((header) => String(header || "").trim().toLowerCase() === normalized);

  if (seriesIndex < 0) {
    return [];
  }

  const rawColumn = seriesIndex + 1;
  const formattedColumn = seriesIndex + 1;

  return plot.rows
    .map((row) => ({
      day: row.raw[0],
      dayRaw: row.formatted[0] || row.source[0],
      value: row.raw[rawColumn],
      valueRaw: row.formatted[formattedColumn] || row.source[formattedColumn],
    }))
    .filter((point) => Number.isFinite(point.day) && Number.isFinite(point.value))
    .sort((a, b) => a.day - b.day);
}

function carriedForwardAccessor(points) {
  let index = 0;
  let current = null;

  return (day) => {
    while (index < points.length && points[index].day <= day + 1e-9) {
      current = points[index];
      index += 1;
    }

    return current;
  };
}

function aggregateCarriedForwardAccessor(pointGroups) {
  const accessors = pointGroups.map((points) => carriedForwardAccessor(points));

  return (day) => {
    let total = 0;
    let found = false;
    let dayRaw = "";

    for (const accessor of accessors) {
      const point = accessor(day);

      if (point && Number.isFinite(point.value)) {
        total += point.value;
        found = true;
        dayRaw ||= point.dayRaw;
      }
    }

    if (!found) {
      return null;
    }

    return {
      day,
      dayRaw: dayRaw || String(day),
      value: total,
      valueRaw: formatMetricNumber(total),
    };
  };
}

function buildBacktestDailyRows(config, plotSnapshots, policySnapshot = []) {
  const servedRegions = servedRegionsForWarehouse(policySnapshot, "warehouse_calopeia");
  const demandRegions = servedRegions.length ? servedRegions : ["Calopeia"];
  const demandGroups = demandRegions.map((region) =>
    seriesPoints(plotSnapshots, "hq_demand", region),
  );
  const lostDemandGroups = demandRegions.map((region) =>
    seriesPoints(plotSnapshots, "hq_lost_demand", region),
  );
  const shipmentGroups = demandRegions.map((region) =>
    seriesPoints(plotSnapshots, "warehouse_shipments", region),
  );
  const sources = {
    inventory: seriesPoints(plotSnapshots, "warehouse_inventory", "warehouse"),
    demand: demandGroups.flat(),
    lostDemand: lostDemandGroups.flat(),
    shipments: shipmentGroups.flat(),
    wip: seriesPoints(plotSnapshots, "factory_wip", "Calopeia"),
    cashBalance: seriesPoints(plotSnapshots, "hq_cash_balance", "value"),
  };
  const days = sortedUniqueNumbers(
    Object.values(sources).flatMap((points) => points.map((point) => point.day)),
  );
  const accessors = Object.fromEntries(
    Object.entries(sources).map(([name, points]) => {
      if (["demand", "lostDemand", "shipments"].includes(name)) {
        const groups =
          name === "demand"
            ? demandGroups
            : name === "lostDemand"
              ? lostDemandGroups
              : shipmentGroups;
        return [name, aggregateCarriedForwardAccessor(groups)];
      }

      return [name, carriedForwardAccessor(points)];
    }),
  );
  const alpha = Number(config.excel?.exponential_smoothing_alpha ?? 0.3);
  const smoothingAlpha = Number.isFinite(alpha) ? alpha : 0.3;
  const rows = [];
  let previousInventory = null;
  let previousCash = null;
  let inventoryEma = null;

  for (const day of days) {
    const points = Object.fromEntries(
      Object.entries(accessors).map(([name, accessor]) => [name, accessor(day)]),
    );
    const inventory = points.inventory?.value ?? null;
    const demand = points.demand?.value ?? null;
    const lostDemand = points.lostDemand?.value ?? null;
    const shipments = points.shipments?.value ?? null;
    const wip = points.wip?.value ?? null;
    const cashBalance = points.cashBalance?.value ?? null;
    const daysOfCover =
      Number.isFinite(inventory) && Number.isFinite(demand) && demand > 0
        ? inventory / demand
        : null;
    const lostDemandRate =
      Number.isFinite(lostDemand) && Number.isFinite(demand) && demand > 0
        ? (lostDemand / demand) * 100
        : null;
    const shipmentToDemandRatio =
      Number.isFinite(shipments) && Number.isFinite(demand) && demand > 0
        ? shipments / demand
        : null;
    const wipToDemandRatio =
      Number.isFinite(wip) && Number.isFinite(demand) && demand > 0
        ? wip / demand
        : null;
    const inventoryDelta =
      Number.isFinite(inventory) && Number.isFinite(previousInventory)
        ? inventory - previousInventory
        : null;
    const cashDelta =
      Number.isFinite(cashBalance) && Number.isFinite(previousCash)
        ? cashBalance - previousCash
        : null;

    if (Number.isFinite(inventory)) {
      inventoryEma =
        inventoryEma === null
          ? inventory
          : smoothingAlpha * inventory + (1 - smoothingAlpha) * inventoryEma;
      previousInventory = inventory;
    }

    if (Number.isFinite(cashBalance)) {
      previousCash = cashBalance;
    }

    rows.push({
      day,
      dayRaw:
        points.inventory?.dayRaw ||
        points.demand?.dayRaw ||
        points.cashBalance?.dayRaw ||
        String(day),
      warehouseInventory: inventory,
      demand,
      lostDemand,
      shipments,
      factoryWip: wip,
      cashBalance,
      inventoryDelta,
      cashDelta,
      daysOfCover,
      lostDemandRate,
      shipmentToDemandRatio,
      wipToDemandRatio,
      inventoryEma,
    });
  }

  return rows;
}

function futureEvent(rows, startIndex, horizonDays, predicate) {
  const startDay = rows[startIndex]?.day;

  if (!Number.isFinite(startDay)) {
    return false;
  }

  for (let index = startIndex; index < rows.length; index += 1) {
    const row = rows[index];

    if (Number.isFinite(row.day) && row.day - startDay > horizonDays) {
      break;
    }

    if (predicate(row)) {
      return true;
    }
  }

  return false;
}

function evaluateBacktestRule(rows, eligiblePredicate, alertPredicate, eventPredicate) {
  let observations = 0;
  let alerts = 0;
  let events = 0;
  let truePositive = 0;
  let falsePositive = 0;
  let falseNegative = 0;
  let trueNegative = 0;

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];

    if (!eligiblePredicate(row, index)) {
      continue;
    }

    const isAlert = Boolean(alertPredicate(row, index));
    const isEvent = Boolean(eventPredicate(row, index));
    observations += 1;
    alerts += isAlert ? 1 : 0;
    events += isEvent ? 1 : 0;

    if (isAlert && isEvent) {
      truePositive += 1;
    } else if (isAlert && !isEvent) {
      falsePositive += 1;
    } else if (!isAlert && isEvent) {
      falseNegative += 1;
    } else {
      trueNegative += 1;
    }
  }

  const precision =
    truePositive + falsePositive > 0
      ? truePositive / (truePositive + falsePositive)
      : null;
  const recall =
    truePositive + falseNegative > 0
      ? truePositive / (truePositive + falseNegative)
      : null;
  const f1 =
    Number.isFinite(precision) && Number.isFinite(recall) && precision + recall > 0
      ? (2 * precision * recall) / (precision + recall)
      : null;

  return {
    observations,
    alerts,
    events,
    truePositive,
    falsePositive,
    falseNegative,
    trueNegative,
    precision,
    recall,
    f1,
  };
}

function backtestCandidateList(values, fallbackValues) {
  const parsed = (Array.isArray(values) ? values : fallbackValues)
    .map(Number)
    .filter(Number.isFinite);

  return sortedUniqueNumbers(parsed.length ? parsed : fallbackValues);
}

function summarizeBacktestRows(rows) {
  const latest = [...rows]
    .reverse()
    .find(
      (row) =>
        Number.isFinite(row.warehouseInventory) ||
        Number.isFinite(row.demand) ||
        Number.isFinite(row.cashBalance),
    );

  if (!latest) {
    return {};
  }

  return {
    day: latest.dayRaw,
    day_number: latest.day,
    warehouse_inventory: latest.warehouseInventory,
    demand: latest.demand,
    lost_demand: latest.lostDemand,
    shipments: latest.shipments,
    factory_wip: latest.factoryWip,
    cash_balance: latest.cashBalance,
    inventory_delta: latest.inventoryDelta,
    cash_delta: latest.cashDelta,
    days_of_cover: latest.daysOfCover,
    lost_demand_rate: latest.lostDemandRate,
    shipment_to_demand_ratio: latest.shipmentToDemandRatio,
    wip_to_demand_ratio: latest.wipToDemandRatio,
    inventory_ema: latest.inventoryEma,
  };
}

function candidateNote(test) {
  const falseNegative = test.falseNegative ?? 0;
  const falsePositive = test.falsePositive ?? 0;

  if (test.events === 0) {
    return "No matching historical event in the current crawl window.";
  }

  if (test.alerts === 0) {
    return "No alerts would have fired at this threshold.";
  }

  if (falseNegative === 0 && falsePositive === 0) {
    return "Matched all events in this historical window.";
  }

  if (falseNegative > falsePositive) {
    return "Missed events are the main weakness.";
  }

  if (falsePositive > falseNegative) {
    return "Extra alerts are the main weakness.";
  }

  return "Balanced misses and extra alerts.";
}

function configuredRuleThreshold(config, ruleId, fallback = null) {
  const rule = (config.monitor?.alert_rules || []).find((item) => item.id === ruleId);
  const value = Number(rule?.threshold);

  return Number.isFinite(value) ? value : fallback;
}

function bestBacktestCandidate(tests, category, preferredCandidate = null) {
  const preferred = Number(preferredCandidate);
  const candidates = tests
    .filter((test) => test.category === category)
    .sort((a, b) => {
      const scoreDiff = (b.f1 ?? -1) - (a.f1 ?? -1);

      if (scoreDiff !== 0) {
        return scoreDiff;
      }

      const missDiff = a.falseNegative - b.falseNegative;

      if (missDiff !== 0) {
        return missDiff;
      }

      const falsePositiveDiff = a.falsePositive - b.falsePositive;

      if (falsePositiveDiff !== 0) {
        return falsePositiveDiff;
      }

      if (Number.isFinite(preferred)) {
        const distanceDiff =
          Math.abs(a.candidate - preferred) - Math.abs(b.candidate - preferred);

        if (distanceDiff !== 0) {
          return distanceDiff;
        }
      }

      return a.candidate - b.candidate;
    });

  return candidates[0] || null;
}

function buildBacktestReport(
  config,
  record,
  standingReport,
  plotSnapshots,
  policySnapshot = [],
) {
  const enabled = config.backtest?.enabled !== false;

  if (!enabled) {
    return {
      enabled: false,
      generated_at: record.checkedAt,
      generated_at_local: record.checkedAtLocal,
      reason: "backtest.enabled is false",
    };
  }

  const rows = buildBacktestDailyRows(config, plotSnapshots, policySnapshot);
  const targets = config.auto_adjust?.targets || {};
  const horizonDays = Number(config.backtest?.horizon_days ?? 3);
  const excessCoverDays = Number(
    config.backtest?.excess_cover_days ?? targets.days_of_cover_max ?? 5,
  );
  const lostDemandThreshold = Number(
    config.backtest?.lost_demand_threshold ?? targets.lost_demand_max ?? 0,
  );
  const currentThreshold = Number(config.monitor?.warehouse_inventory_threshold);
  const inventoryCandidates = backtestCandidateList(
    config.backtest?.inventory_threshold_candidates,
    [250, 350, currentThreshold, 600].filter(Number.isFinite),
  );
  const coverHighCandidates = backtestCandidateList(
    config.backtest?.days_of_cover_high_candidates,
    [3, 5, 7, 10],
  );
  const coverLowCandidates = backtestCandidateList(
    config.backtest?.days_of_cover_low_candidates,
    [0.5, 1, 2],
  );
  const shipmentRatioCandidates = backtestCandidateList(
    config.backtest?.shipment_to_demand_ratio_candidates,
    [0.6, 0.8, 0.9, 1],
  );
  const highCoverEvent = (row) =>
    Number.isFinite(row.daysOfCover) && row.daysOfCover > excessCoverDays;
  const shortageEvent = (_, index) =>
    futureEvent(
      rows,
      index,
      Number.isFinite(horizonDays) ? horizonDays : 3,
      (row) =>
        Number.isFinite(row.lostDemand) && row.lostDemand > lostDemandThreshold,
    );
  const tests = [];

  for (const threshold of inventoryCandidates) {
    const result = evaluateBacktestRule(
      rows,
      (row) => Number.isFinite(row.warehouseInventory) && Number.isFinite(row.daysOfCover),
      (row) => row.warehouseInventory >= threshold,
      highCoverEvent,
    );

    tests.push({
      category: "warehouse_inventory_high",
      indicator: "Warehouse inventory high threshold",
      candidate: threshold,
      operator: ">=",
      event_definition: `days_of_cover > ${excessCoverDays}`,
      ...result,
    });
  }

  for (const threshold of coverHighCandidates) {
    const result = evaluateBacktestRule(
      rows,
      (row) => Number.isFinite(row.daysOfCover),
      (row) => row.daysOfCover >= threshold,
      highCoverEvent,
    );

    tests.push({
      category: "days_of_cover_high",
      indicator: "Days of cover high threshold",
      candidate: threshold,
      operator: ">=",
      event_definition: `days_of_cover > ${excessCoverDays}`,
      ...result,
    });
  }

  for (const threshold of coverLowCandidates) {
    const result = evaluateBacktestRule(
      rows,
      (row, index) => Number.isFinite(row.daysOfCover) && rows.length > index,
      (row) => row.daysOfCover <= threshold,
      shortageEvent,
    );

    tests.push({
      category: "days_of_cover_low",
      indicator: "Days of cover low threshold",
      candidate: threshold,
      operator: "<=",
      event_definition: `lost_demand > ${lostDemandThreshold} within ${horizonDays} days`,
      ...result,
    });
  }

  for (const threshold of shipmentRatioCandidates) {
    const result = evaluateBacktestRule(
      rows,
      (row) => Number.isFinite(row.shipmentToDemandRatio),
      (row) => row.shipmentToDemandRatio < threshold,
      shortageEvent,
    );

    tests.push({
      category: "shipment_to_demand_ratio_low",
      indicator: "Shipment / demand ratio low threshold",
      candidate: threshold,
      operator: "<",
      event_definition: `lost_demand > ${lostDemandThreshold} within ${horizonDays} days`,
      ...result,
    });
  }

  for (const test of tests) {
    test.precision_text = Number.isFinite(test.precision)
      ? percentMetric(test.precision * 100)
      : "";
    test.recall_text = Number.isFinite(test.recall)
      ? percentMetric(test.recall * 100)
      : "";
    test.f1_text = Number.isFinite(test.f1) ? ratioMetric(test.f1) : "";
    test.note = candidateNote(test);
  }

  const recommendations = [
    bestBacktestCandidate(tests, "warehouse_inventory_high", currentThreshold),
    bestBacktestCandidate(
      tests,
      "days_of_cover_high",
      configuredRuleThreshold(config, "days_of_cover_high", targets.days_of_cover_max),
    ),
    bestBacktestCandidate(
      tests,
      "days_of_cover_low",
      configuredRuleThreshold(config, "days_of_cover_low", targets.days_of_cover_min),
    ),
    bestBacktestCandidate(
      tests,
      "shipment_to_demand_ratio_low",
      configuredRuleThreshold(
        config,
        "shipments_below_demand",
        targets.shipment_to_demand_ratio_min,
      ),
    ),
  ]
    .filter(Boolean)
    .map((test) => ({
      category: test.category,
      indicator: test.indicator,
      suggested_threshold: test.candidate,
      operator: test.operator,
      event_definition: test.event_definition,
      f1: roundMetric(test.f1, 4),
      precision: roundMetric(test.precision, 4),
      recall: roundMetric(test.recall, 4),
      note: test.note,
    }));

  return {
    enabled: true,
    generated_at: record.checkedAt,
    generated_at_local: record.checkedAtLocal,
    target_team: record.targetTeam,
    target_rank: record.targetRank,
    target_cash: record.targetCash,
    dashboard_day: record.dashboardDay,
    data_source: config.backtest?.data_source || "current trigger crawl plot data",
    local_dependency: "none",
    horizon_days: Number.isFinite(horizonDays) ? horizonDays : 3,
    excess_cover_days: Number.isFinite(excessCoverDays) ? excessCoverDays : 5,
    lost_demand_threshold: Number.isFinite(lostDemandThreshold)
      ? lostDemandThreshold
      : 0,
    observations: rows.length,
    day_start: backtestDayLabel(rows[0]?.day, rows[0]?.dayRaw || ""),
    day_end: backtestDayLabel(
      rows[rows.length - 1]?.day,
      rows[rows.length - 1]?.dayRaw || "",
    ),
    latest: summarizeBacktestRows(rows),
    recommendations,
    tests,
    daily_rows: rows,
    standing_reference: {
      team: standingReport.target.team,
      rank: standingReport.target.rank,
      cash: standingReport.target.cash,
    },
  };
}

function buildBacktestSummaryCsv(report) {
  const header = [
    "category",
    "indicator",
    "operator",
    "candidate",
    "event_definition",
    "observations",
    "alerts",
    "events",
    "true_positive",
    "false_positive",
    "false_negative",
    "true_negative",
    "precision",
    "recall",
    "f1",
    "note",
  ];
  const rows = (report?.tests || []).map((test) => [
    test.category,
    test.indicator,
    test.operator,
    test.candidate,
    test.event_definition,
    test.observations,
    test.alerts,
    test.events,
    test.truePositive,
    test.falsePositive,
    test.falseNegative,
    test.trueNegative,
    roundMetric(test.precision, 4),
    roundMetric(test.recall, 4),
    roundMetric(test.f1, 4),
    test.note,
  ]);

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
}

function buildBacktestDailyCsv(report) {
  const header = [
    "day",
    "warehouse_inventory",
    "demand",
    "lost_demand",
    "shipments",
    "factory_wip",
    "cash_balance",
    "inventory_delta",
    "cash_delta",
    "days_of_cover",
    "lost_demand_rate_percent",
    "shipment_to_demand_ratio",
    "wip_to_demand_ratio",
    "inventory_ema",
  ];
  const rows = (report?.daily_rows || []).map((row) => [
    row.dayRaw,
    row.warehouseInventory,
    row.demand,
    row.lostDemand,
    row.shipments,
    row.factoryWip,
    row.cashBalance,
    row.inventoryDelta,
    row.cashDelta,
    row.daysOfCover,
    row.lostDemandRate,
    row.shipmentToDemandRatio,
    row.wipToDemandRatio,
    row.inventoryEma,
  ]);

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map((value) => csvValue(value ?? "")).join(",")),
  ].join("\r\n")}\r\n`;
}

function styleSectionHeader(row, fill = "FFE2E8F0", fontColor = "FF0F172A") {
  row.font = { bold: true, color: { argb: fontColor } };
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: fill },
  };
  row.alignment = { vertical: "middle" };
}

function addBacktestDashboardWorksheet(workbook, usedNames, report) {
  if (!report?.enabled) {
    return;
  }

  const worksheet = workbook.addWorksheet(safeWorksheetName("Backtest", usedNames));
  const latest = report.latest || {};

  worksheet.views = [{ state: "frozen", ySplit: 3 }];
  worksheet.columns = [
    { width: 30 },
    { width: 18 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 44 },
  ];

  worksheet.mergeCells("A1:G1");
  const titleCell = worksheet.getCell("A1");
  titleCell.value = "Backtest";
  titleCell.font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
  titleCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1D4ED8" },
  };
  titleCell.alignment = { vertical: "middle" };
  worksheet.getRow(1).height = 28;

  worksheet.addRow([
    "Generated at",
    report.generated_at_local,
    "Data source",
    report.data_source,
    "Local dependency",
    report.local_dependency,
    "",
  ]);
  worksheet.addRow([
    "Window",
    `${report.day_start} to ${report.day_end}`,
    "Observations",
    report.observations,
    "Horizon days",
    report.horizon_days,
    "",
  ]);

  const latestHeader = worksheet.addRow(["Latest Operational State", "", "", "", "", "", ""]);
  worksheet.mergeCells(`A${latestHeader.number}:G${latestHeader.number}`);
  styleSectionHeader(latestHeader, "FFDDEBFF");
  worksheet.addRow([
    "warehouse_inventory",
    latest.warehouse_inventory ?? "",
    "demand",
    latest.demand ?? "",
    "lost_demand",
    latest.lost_demand ?? "",
    "",
  ]);
  worksheet.addRow([
    "shipments",
    latest.shipments ?? "",
    "factory_wip",
    latest.factory_wip ?? "",
    "cash_balance",
    latest.cash_balance ?? "",
    "",
  ]);
  worksheet.addRow([
    "days_of_cover",
    roundMetric(latest.days_of_cover, 4) ?? "",
    "shipment_to_demand_ratio",
    roundMetric(latest.shipment_to_demand_ratio, 4) ?? "",
    "inventory_ema",
    roundMetric(latest.inventory_ema, 4) ?? "",
    "",
  ]);

  worksheet.addRow([]);
  const recHeader = worksheet.addRow([
    "Recommended indicator",
    "Operator",
    "Suggested threshold",
    "Precision",
    "Recall",
    "F1",
    "Basis",
  ]);
  styleSectionHeader(recHeader);

  for (const item of report.recommendations || []) {
    worksheet.addRow([
      item.indicator,
      item.operator,
      item.suggested_threshold,
      Number.isFinite(item.precision) ? item.precision : "",
      Number.isFinite(item.recall) ? item.recall : "",
      Number.isFinite(item.f1) ? item.f1 : "",
      item.note,
    ]);
  }

  worksheet.addRow([]);
  const thresholdHeader = worksheet.addRow([
    "category",
    "candidate",
    "operator",
    "observations",
    "precision",
    "recall",
    "f1",
  ]);
  styleSectionHeader(thresholdHeader);

  for (const test of report.tests || []) {
    worksheet.addRow([
      test.category,
      test.candidate,
      test.operator,
      test.observations,
      roundMetric(test.precision, 4) ?? "",
      roundMetric(test.recall, 4) ?? "",
      roundMetric(test.f1, 4) ?? "",
    ]);
  }

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE2E8F0" } },
        left: { style: "thin", color: { argb: "FFE2E8F0" } },
        bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
        right: { style: "thin", color: { argb: "FFE2E8F0" } },
      };
      cell.alignment = { vertical: "middle", wrapText: true };
    });
  });
}

function addBacktestSummaryWorksheet(workbook, usedNames, report) {
  if (!report?.enabled) {
    return;
  }

  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Backtest Summary", usedNames),
  );
  const latest = report.latest || {};

  worksheet.addRow(["Field", "Value"]);
  worksheet.addRow(["Generated at", report.generated_at_local]);
  worksheet.addRow(["Data source", report.data_source]);
  worksheet.addRow(["Local dependency", report.local_dependency]);
  worksheet.addRow(["Target team", report.target_team]);
  worksheet.addRow(["Target rank", report.target_rank]);
  worksheet.addRow(["Target cash", report.target_cash]);
  worksheet.addRow(["Dashboard day", report.dashboard_day]);
  worksheet.addRow(["Observation days", `${report.day_start} to ${report.day_end}`]);
  worksheet.addRow(["Observation count", report.observations]);
  worksheet.addRow(["Latest inventory", latest.warehouse_inventory ?? ""]);
  worksheet.addRow(["Latest demand", latest.demand ?? ""]);
  worksheet.addRow(["Latest days of cover", roundMetric(latest.days_of_cover, 4) ?? ""]);
  worksheet.addRow([
    "Latest shipment / demand",
    roundMetric(latest.shipment_to_demand_ratio, 4) ?? "",
  ]);
  worksheet.addRow(["Latest lost demand", latest.lost_demand ?? ""]);
  worksheet.addRow(["Latest cash balance", latest.cash_balance ?? ""]);
  worksheet.addRow(["Horizon days", report.horizon_days]);
  worksheet.addRow(["Excess cover event", `days_of_cover > ${report.excess_cover_days}`]);
  worksheet.addRow([
    "Shortage event",
    `lost_demand > ${report.lost_demand_threshold} within ${report.horizon_days} days`,
  ]);

  worksheet.addRow([]);
  worksheet.addRow([
    "Recommended indicator",
    "Suggested threshold",
    "Operator",
    "Precision",
    "Recall",
    "F1",
    "Basis",
  ]);

  for (const recommendation of report.recommendations || []) {
    worksheet.addRow([
      recommendation.indicator,
      recommendation.suggested_threshold,
      recommendation.operator,
      recommendation.precision,
      recommendation.recall,
      recommendation.f1,
      recommendation.note,
    ]);
  }

  styleWorksheet(worksheet);
}

function addBacktestThresholdWorksheet(workbook, usedNames, report) {
  if (!report?.enabled || !report.tests?.length) {
    return;
  }

  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Backtest Thresholds", usedNames),
  );

  worksheet.addRow([
    "category",
    "indicator",
    "operator",
    "candidate",
    "event_definition",
    "observations",
    "alerts",
    "events",
    "true_positive",
    "false_positive",
    "false_negative",
    "true_negative",
    "precision",
    "recall",
    "f1",
    "note",
  ]);

  for (const test of report.tests) {
    worksheet.addRow([
      test.category,
      test.indicator,
      test.operator,
      test.candidate,
      test.event_definition,
      test.observations,
      test.alerts,
      test.events,
      test.truePositive,
      test.falsePositive,
      test.falseNegative,
      test.trueNegative,
      roundMetric(test.precision, 4) ?? "",
      roundMetric(test.recall, 4) ?? "",
      roundMetric(test.f1, 4) ?? "",
      test.note,
    ]);
  }

  styleWorksheet(worksheet);
}

function addBacktestDailyWorksheet(workbook, usedNames, report) {
  if (!report?.enabled || !report.daily_rows?.length) {
    return;
  }

  const worksheet = workbook.addWorksheet(safeWorksheetName("Backtest Daily", usedNames));

  worksheet.addRow([
    "day",
    "warehouse_inventory",
    "demand",
    "lost_demand",
    "shipments",
    "factory_wip",
    "cash_balance",
    "inventory_delta",
    "cash_delta",
    "days_of_cover",
    "lost_demand_rate_percent",
    "shipment_to_demand_ratio",
    "wip_to_demand_ratio",
    "inventory_ema",
  ]);

  for (const row of report.daily_rows) {
    worksheet.addRow([
      row.dayRaw,
      row.warehouseInventory ?? "",
      row.demand ?? "",
      row.lostDemand ?? "",
      row.shipments ?? "",
      row.factoryWip ?? "",
      row.cashBalance ?? "",
      row.inventoryDelta ?? "",
      row.cashDelta ?? "",
      roundMetric(row.daysOfCover, 4) ?? "",
      roundMetric(row.lostDemandRate, 4) ?? "",
      roundMetric(row.shipmentToDemandRatio, 4) ?? "",
      roundMetric(row.wipToDemandRatio, 4) ?? "",
      roundMetric(row.inventoryEma, 4) ?? "",
    ]);
  }

  styleWorksheet(worksheet);
}

function addBacktestWorksheets(workbook, usedNames, report) {
  addBacktestDashboardWorksheet(workbook, usedNames, report);
  addBacktestSummaryWorksheet(workbook, usedNames, report);
  addBacktestThresholdWorksheet(workbook, usedNames, report);
  addBacktestDailyWorksheet(workbook, usedNames, report);
}

async function buildBacktestWorkbookBuffer(report) {
  if (!report?.enabled) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "MGT267 Watchdog";
  workbook.created = new Date(report.generated_at);
  workbook.modified = new Date(report.generated_at);
  workbook.calcProperties.fullCalcOnLoad = true;
  const usedNames = new Set();

  addBacktestWorksheets(workbook, usedNames, report);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
}

async function writeBacktestOutputs(config, report) {
  const jsonPath = path.resolve(
    process.cwd(),
    config.output.backtest_json || ".monitor-state/backtest_latest.json",
  );
  const summaryCsvPath = path.resolve(
    process.cwd(),
    config.output.backtest_summary_csv ||
      ".monitor-state/backtest_summary_latest.csv",
  );
  const dailyCsvPath = path.resolve(
    process.cwd(),
    config.output.backtest_daily_csv || ".monitor-state/backtest_daily_latest.csv",
  );
  const workbookPath = path.resolve(
    process.cwd(),
    config.output.backtest_xlsx || ".monitor-state/backtest_report_latest.xlsx",
  );

  ensureDir(path.dirname(jsonPath));
  fs.writeFileSync(jsonPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  fs.writeFileSync(summaryCsvPath, buildBacktestSummaryCsv(report), "utf8");
  fs.writeFileSync(dailyCsvPath, buildBacktestDailyCsv(report), "utf8");

  const buffer = await buildBacktestWorkbookBuffer(report);

  if (buffer) {
    fs.writeFileSync(workbookPath, buffer);
  }

  return {
    jsonPath,
    summaryCsvPath,
    dailyCsvPath,
    workbookPath: buffer ? workbookPath : "",
  };
}

function buildPolicySnapshotCsv(snapshot = []) {
  const header = [
    "page_id",
    "section",
    "region",
    "page_label",
    "item_type",
    "form_index",
    "table_index",
    "row_index",
    "column_index",
    "name",
    "type",
    "value",
    "text",
    "checked",
  ];
  const rows = [];

  for (const page of snapshot) {
    for (const fact of page.facts || []) {
      rows.push([
        page.id,
        page.section,
        page.region,
        page.label,
        "fact",
        "",
        "",
        "",
        "",
        "",
        "",
        fact,
        "",
        "",
      ]);
    }

    for (const control of page.forms || []) {
      rows.push([
        page.id,
        page.section,
        page.region,
        page.label,
        "form_control",
        control.formIndex,
        "",
        "",
        "",
        control.name,
        control.type,
        control.value,
        control.text,
        control.checked,
      ]);
    }

    for (const table of page.tables || []) {
      for (const [rowIndex, row] of table.rows.entries()) {
        for (const [columnIndex, value] of row.entries()) {
          rows.push([
            page.id,
            page.section,
            page.region,
            page.label,
            "table_cell",
            "",
            table.index,
            rowIndex + 1,
            columnIndex + 1,
            table.label,
            "",
            value,
            "",
            "",
          ]);
        }
      }
    }
  }

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
}

function findPolicyPage(policySnapshot, id) {
  return (policySnapshot || []).find((page) => page.id === id) || null;
}

function findPolicyControl(page, name) {
  return (page?.forms || []).find((control) => control.name === name) || null;
}

function numericPolicyValue(control) {
  if (!control) {
    return undefined;
  }

  const value = Number(control.value);
  return Number.isFinite(value) ? value : undefined;
}

function buildScrapedPolicyBaseline(policySnapshot) {
  const factoryPage = findPolicyPage(policySnapshot, "factory_calopeia");
  const warehousePage = findPolicyPage(policySnapshot, "warehouse_calopeia");
  const factoryShip = findPolicyControl(factoryPage, "ship1");
  const warehouseShip = findPolicyControl(warehousePage, "ship1");

  return {
    factory: {
      shipping_method: factoryShip?.value,
      shipping_method_text: factoryShip?.text,
      order_point: numericPolicyValue(findPolicyControl(factoryPage, "point1")),
      quantity: numericPolicyValue(findPolicyControl(factoryPage, "quant1")),
      priority: numericPolicyValue(findPolicyControl(factoryPage, "priority1")),
    },
    warehouse: {
      shipping_method: warehouseShip?.value,
      shipping_method_text: warehouseShip?.text,
      order_point: numericPolicyValue(findPolicyControl(warehousePage, "point1")),
      quantity: numericPolicyValue(findPolicyControl(warehousePage, "quant1")),
      priority: numericPolicyValue(findPolicyControl(warehousePage, "priority1")),
    },
  };
}

function mergePolicyBaseline(fallbackBaseline = {}, scrapedBaseline = {}) {
  return {
    factory: {
      ...(fallbackBaseline.factory || {}),
      ...Object.fromEntries(
        Object.entries(scrapedBaseline.factory || {}).filter(
          ([, value]) => value !== undefined && value !== "",
        ),
      ),
    },
    warehouse: {
      ...(fallbackBaseline.warehouse || {}),
      ...Object.fromEntries(
        Object.entries(scrapedBaseline.warehouse || {}).filter(
          ([, value]) => value !== undefined && value !== "",
        ),
      ),
    },
  };
}

function gameRules(config) {
  return {
    ...DEFAULT_GAME_RULES,
    ...(config.game_rules || {}),
  };
}

function dashboardDayNumber(record) {
  const value = Number(String(record.dashboardDay ?? "").replace(/,/g, ""));
  return Number.isFinite(value) ? value : null;
}

function remainingGameDays(config, record) {
  const day = dashboardDayNumber(record);
  const endDay = Number(gameRules(config).end_day);

  if (!Number.isFinite(day) || !Number.isFinite(endDay)) {
    return null;
  }

  return Math.max(0, endDay - day);
}

function coverageTargets(config, record) {
  const targets = config.auto_adjust?.targets || {};
  const baseMin = Number(targets.days_of_cover_min ?? 2);
  const baseMax = Number(targets.days_of_cover_max ?? 5);
  const baseTarget = Number(
    targets.days_of_cover_target ?? (baseMin + baseMax) / 2,
  );
  const remainingDays = remainingGameDays(config, record);
  const maxByGameClock = Number.isFinite(remainingDays)
    ? Math.max(0, remainingDays)
    : baseMax;
  const max = Math.max(0, Math.min(baseMax, maxByGameClock));
  const target = Math.max(0, Math.min(baseTarget, max));
  const min = Math.max(0, Math.min(baseMin, target || max));

  return {
    min,
    max,
    target,
    baseMin,
    baseMax,
    baseTarget,
    remainingDays,
  };
}

function policyBaseline(config, policySnapshot) {
  return mergePolicyBaseline(
    config.auto_adjust?.policy_baseline || {},
    buildScrapedPolicyBaseline(policySnapshot || []),
  );
}

function servedRegionsForWarehouse(policySnapshot, pageId = "warehouse_calopeia") {
  const page = findPolicyPage(policySnapshot, pageId);
  const outboundTable = (page?.tables || []).find((table) =>
    /outbound shipments to customers/i.test(table.label || ""),
  );
  const regions = [];

  for (const row of (outboundTable?.rows || []).slice(1)) {
    const destination = compactText(row[0] || "");
    const served = compactText(row[2] || "").toLowerCase();

    if (destination && served.startsWith("yes")) {
      regions.push(destination);
    }
  }

  return [...new Set(regions)];
}

function metricMapValue(metricMap, key) {
  const metric = metricMap.get(key);
  return Number.isFinite(metric?.valueNumber) ? metric.valueNumber : null;
}

function sumMetricForRegions(metricMap, prefix, regions) {
  let total = 0;
  let found = false;

  for (const region of regions) {
    const value = metricMapValue(metricMap, `${prefix}:${region}`);

    if (Number.isFinite(value)) {
      total += value;
      found = true;
    }
  }

  return found ? total : null;
}

function policyMethodText(policy = {}) {
  return policy.shipping_method_text || policy.shipping_method || "n/a";
}

function policyNumberText(value) {
  return value === undefined || value === "" ? "n/a" : String(value);
}

function policyContextLines(config, policySnapshot) {
  const baseline = policyBaseline(config, policySnapshot);
  const factory = baseline.factory || {};
  const warehouse = baseline.warehouse || {};

  return [
    `Factory policy: shipping_method=${policyMethodText(factory)}, order_point=${policyNumberText(factory.order_point)}, quantity=${policyNumberText(factory.quantity)}, priority=${policyNumberText(factory.priority)}.`,
    `Warehouse policy: shipping_method=${policyMethodText(warehouse)}, order_point=${policyNumberText(warehouse.order_point)}, quantity=${policyNumberText(warehouse.quantity)}, priority=${policyNumberText(warehouse.priority)}.`,
  ];
}

function gameRuleContextLines(config, record) {
  const rules = gameRules(config);
  const remainingDays = remainingGameDays(config, record);
  const truckBreakEven =
    Number(rules.mail_cost_per_drum) > 0
      ? Number(rules.truck_cost) / Number(rules.mail_cost_per_drum)
      : null;

  return [
    `Game ends on day ${rules.end_day}; current dashboard day is ${record.dashboardDay}; remaining days are ${remainingDays === null ? "n/a" : remainingDays}. Inventory and capacity are obsolete at game end.`,
    `Orders must ship within ${rules.order_response_hours} hours or become lost demand; future demand is otherwise stable and seasonal.`,
    `Economics: price=${rules.product_price_per_drum}/drum, customer fulfillment cost=${rules.customer_fulfillment_cost_per_drum}/drum, holding cost=${rules.holding_cost_per_drum_per_year}/drum/year, cash earns ${rules.cash_interest_rate_annual_percent}%/year.`,
    `Capacity expansion costs ${rules.capacity_expansion_cost_per_drum_per_day} per added drum/day, takes ${rules.capacity_expansion_lead_days} days, and cannot be retired.`,
    `Production batch cost is ${rules.production_fixed_cost_per_batch} fixed plus ${rules.production_variable_cost_per_drum}/drum; avoid tiny batches unless flexibility is worth the fixed cost.`,
    `Factory-to-warehouse transport: mail costs ${rules.mail_cost_per_drum}/drum and takes ${rules.mail_lead_days} day; truck costs ${rules.truck_cost} up to ${rules.truck_capacity_drums} drums and takes ${rules.truck_lead_days} days${Number.isFinite(truckBreakEven) ? `, so truck is cheaper than mail only above about ${Math.ceil(truckBreakEven)} drums` : ""}.`,
    "Priority level does not affect this assignment; never recommend changing priority.",
  ];
}

function serviceContextLines(config, record, metricCatalog, policySnapshot) {
  const regions = servedRegionsForWarehouse(policySnapshot || [], "warehouse_calopeia");
  const servedRegions = regions.length ? regions : ["Calopeia"];
  const targets = coverageTargets(config, record);
  const demand =
    metricRaw(metricCatalog, "derived:calopeia_served_demand") ||
    metricRaw(metricCatalog, "hq_demand:Calopeia") ||
    "n/a";
  const lostDemand =
    metricRaw(metricCatalog, "derived:calopeia_served_lost_demand") ||
    metricRaw(metricCatalog, "hq_lost_demand:Calopeia") ||
    "n/a";
  const shipments =
    metricRaw(metricCatalog, "derived:calopeia_served_shipments") ||
    metricRaw(metricCatalog, "warehouse_shipments:Calopeia") ||
    "n/a";
  const daysCover = metricValue(metricCatalog, "derived:days_of_cover");
  const daysCoverText = Number.isFinite(daysCover)
    ? formatMetricNumber(daysCover, "days")
    : "n/a";

  return [
    `Calopeia warehouse currently serves ${servedRegions.length} region(s): ${servedRegions.join(", ")}.`,
    `Use served-region demand for Calopeia buffer decisions: served demand=${demand}, served shipments=${shipments}, served lost demand=${lostDemand}, served days of cover=${daysCoverText}.`,
    "The legacy 450 inventory number is not an active alert threshold; it has no standalone basis without served demand and remaining-day context.",
    `If served days of cover is between ${formatMetricNumber(targets.min, "days")} and ${formatMetricNumber(targets.max, "days")} and lost demand is zero, recommend holding or restoring policy settings rather than reducing inventory.`,
    "Avoid bullwhip over-correction: do not swing order point or quantity based on one checkpoint; prefer small staged changes and verify truck/mail/WIP pipeline unless lost demand appears.",
  ];
}

function metricValue(metricCatalog, key) {
  const metric = metricCatalog?.get(key);
  return Number.isFinite(metric?.valueNumber) ? metric.valueNumber : null;
}

function metricRaw(metricCatalog, key) {
  const metric = metricCatalog?.get(key);
  return metric?.valueRaw || "";
}

function clamp(value, min, max) {
  let next = value;

  if (Number.isFinite(min)) {
    next = Math.max(min, next);
  }

  if (Number.isFinite(max)) {
    next = Math.min(max, next);
  }

  return next;
}

function roundedPolicyValue(value) {
  if (!Number.isFinite(value)) {
    return null;
  }

  return Math.max(0, Math.round(value));
}

function suggestPolicyNumber(current, direction, maxChange, min, max) {
  if (!Number.isFinite(current)) {
    return {
      baseline: "",
      suggested: "",
      change: "",
    };
  }

  if (direction === "hold") {
    return {
      baseline: String(current),
      suggested: String(current),
      change: "0",
    };
  }

  const signedChange = direction === "increase" ? maxChange : -maxChange;
  const suggested = roundedPolicyValue(clamp(current + signedChange, min, max));
  const change = suggested - current;

  return {
    baseline: String(current),
    suggested: String(suggested),
    change: formatDelta(change),
  };
}

function buildAutoAdjustmentPlan(
  config,
  metricCatalog,
  record,
  standingReport,
  policySnapshot = [],
) {
  const cfg = config.auto_adjust || {};
  const enabled = cfg.enabled !== false;
  const mode = cfg.mode || "research_only";
  const targets = cfg.targets || {};
  const maxChange = cfg.max_change_per_run || {};
  const bounds = cfg.bounds || {};
  const scrapedBaseline = buildScrapedPolicyBaseline(policySnapshot);
  const baseline = policyBaseline(config, policySnapshot);
  const recommendations = [];

  const demand =
    metricValue(metricCatalog, "derived:calopeia_served_demand") ??
    metricValue(metricCatalog, "hq_demand:Calopeia");
  const shipments =
    metricValue(metricCatalog, "derived:calopeia_served_shipments") ??
    metricValue(metricCatalog, "warehouse_shipments:Calopeia");
  const lostDemand =
    metricValue(metricCatalog, "derived:calopeia_served_lost_demand") ??
    metricValue(metricCatalog, "hq_lost_demand:Calopeia");
  const wip = metricValue(metricCatalog, "factory_wip:Calopeia");
  const daysOfCover = metricValue(metricCatalog, "derived:days_of_cover");
  const shipmentRatio = metricValue(
    metricCatalog,
    "derived:shipment_to_demand_ratio",
  );
  const lostDemandRate = metricValue(metricCatalog, "derived:lost_demand_rate");
  const cashLead = metricValue(
    metricCatalog,
    "derived:cash_lead_percent_vs_nearest",
  );
  const servedRegionCount =
    metricValue(metricCatalog, "derived:calopeia_served_region_count") ?? 1;
  const truckPipeline = metricValue(metricCatalog, "warehouse_inventory:truck");
  const mailPipeline = metricValue(metricCatalog, "warehouse_inventory:mail");
  const rules = gameRules(config);

  const coverTargets = coverageTargets(config, record);
  const daysMin = coverTargets.min;
  const daysMax = coverTargets.max;
  const daysTarget = coverTargets.target;
  const inventoryLow = Number(targets.warehouse_inventory_low ?? 50);
  const inventoryHigh = Number(targets.warehouse_inventory_high);
  const lostDemandMax = Number(targets.lost_demand_max ?? 0);
  const shipmentRatioMin = Number(targets.shipment_to_demand_ratio_min ?? 0.9);
  const wipRatioMax = Number(targets.wip_to_demand_ratio_max ?? 3);
  const orderPointStep = Number(maxChange.order_point ?? 25);
  const quantityStep = Number(maxChange.quantity ?? 25);
  const pointMin = Number(bounds.order_point_min ?? 0);
  const pointMax = Number(bounds.order_point_max ?? 999999);
  const quantityMin = Number(bounds.quantity_min ?? 0);
  const quantityMax = Number(bounds.quantity_max ?? 999999);
  const factory = baseline.factory || {};
  const warehouse = baseline.warehouse || {};
  const targetInventory =
    Number.isFinite(demand) && demand > 0 && Number.isFinite(daysTarget)
      ? roundedPolicyValue(demand * daysTarget)
      : null;
  const excessCoverage =
    (Number.isFinite(daysOfCover) && daysOfCover > daysMax) ||
    (servedRegionCount <= 1 &&
      Number.isFinite(inventoryHigh) &&
      record.warehouseInventory >= inventoryHigh);
  const shortageRisk =
    (Number.isFinite(lostDemand) && lostDemand > lostDemandMax) ||
    (Number.isFinite(daysOfCover) && daysOfCover < daysMin) ||
    record.warehouseInventory <= inventoryLow;
  const shipmentBelowDemand =
    Number.isFinite(shipmentRatio) && shipmentRatio < shipmentRatioMin;
  const wipRatio =
    Number.isFinite(wip) && Number.isFinite(demand) && demand > 0
      ? wip / demand
      : null;
  const activePipeline =
    (Number.isFinite(wip) && wip > 0) ||
    (Number.isFinite(truckPipeline) && truckPipeline > 0) ||
    (Number.isFinite(mailPipeline) && mailPipeline > 0);

  function addRecommendation({
    area,
    parameter,
    baselineValue,
    suggestedValue,
    direction,
    urgency,
    confidence,
    reason,
    change = "",
  }) {
    recommendations.push({
      area,
      parameter,
      baseline: baselineValue === undefined ? "" : String(baselineValue),
      suggested: suggestedValue === undefined ? "" : String(suggestedValue),
      change,
      direction,
      urgency,
      confidence,
      reason,
      submit_allowed: false,
    });
  }

  function addNumericPolicy(area, parameter, current, direction, step, reason) {
    const suggested = suggestPolicyNumber(
      Number(current),
      direction,
      step,
      parameter === "quantity" ? quantityMin : pointMin,
      parameter === "quantity" ? quantityMax : pointMax,
    );
    addRecommendation({
      area,
      parameter,
      baselineValue: suggested.baseline,
      suggestedValue: suggested.suggested,
      change: suggested.change,
      direction,
      urgency: direction === "hold" ? "low" : shortageRisk ? "high" : "medium",
      confidence: Number.isFinite(current) ? "medium" : "low",
      reason,
    });
  }

  function shortageDirection(current, parameter) {
    const numericCurrent = Number(current);

    if (!Number.isFinite(numericCurrent) || targetInventory === null) {
      return "increase";
    }

    if (parameter === "quantity" && numericCurrent >= targetInventory) {
      return "hold";
    }

    if (parameter === "order_point" && numericCurrent >= targetInventory) {
      return "hold";
    }

    return "increase";
  }

  if (!enabled) {
    return {
      enabled: false,
      mode,
      generated_at: record.checkedAt,
      generated_at_local: record.checkedAtLocal,
      safety: {
        research_only: true,
        game_updates_enabled: false,
        submit_allowed: false,
        note: "Auto-adjustment research is disabled; no game-setting changes are possible.",
      },
      inputs: {},
      targets,
      recommendations,
    };
  }

  const posture = shortageRisk ? "shortage_risk" : excessCoverage ? "excess_stock" : "balanced";
  const inventoryReason = [
    `inventory ${record.warehouseInventory}`,
    Number.isFinite(daysOfCover) ? `days of cover ${formatMetricNumber(daysOfCover, "days")}` : "",
    Number.isFinite(demand)
      ? `served demand ${metricRaw(metricCatalog, "derived:calopeia_served_demand") || metricRaw(metricCatalog, "hq_demand:Calopeia")}`
      : "",
    Number.isFinite(lostDemand)
      ? `served lost demand ${metricRaw(metricCatalog, "derived:calopeia_served_lost_demand") || metricRaw(metricCatalog, "hq_lost_demand:Calopeia")}`
      : "",
    targetInventory !== null ? `target inventory near ${targetInventory}` : "",
  ]
    .filter(Boolean)
    .join("; ");

  if (posture === "shortage_risk") {
    const factoryOrderPointDirection = shortageDirection(
      factory.order_point,
      "order_point",
    );
    const factoryQuantityDirection = shortageDirection(factory.quantity, "quantity");
    const warehouseOrderPointDirection = shortageDirection(
      warehouse.order_point,
      "order_point",
    );
    const warehouseQuantityDirection = shortageDirection(
      warehouse.quantity,
      "quantity",
    );
    const pipelineText = activePipeline
      ? `active pipeline detected: WIP ${formatMetricNumber(wip || 0)}, truck ${formatMetricNumber(truckPipeline || 0)}, mail ${formatMetricNumber(mailPipeline || 0)}`
      : "no active pipeline detected";

    addNumericPolicy(
      "factory",
      "order_point",
      factory.order_point,
      factoryOrderPointDirection,
      orderPointStep,
      factoryOrderPointDirection === "hold"
        ? `Shortage risk is timing-driven, but factory order point already exceeds served-demand target; ${inventoryReason}; ${pipelineText}.`
        : `Shortage risk detected and factory order point is below served-demand target; ${inventoryReason}.`,
    );
    addNumericPolicy(
      "factory",
      "quantity",
      factory.quantity,
      factoryQuantityDirection,
      quantityStep,
      factoryQuantityDirection === "hold"
        ? `Current batch quantity is already large relative to served demand; avoid bullwhip and wait for pipeline unless lost demand persists.`
        : `Raise replenishment cautiously until coverage returns to ${daysMin}-${daysMax} days.`,
    );
    addNumericPolicy(
      "warehouse",
      "order_point",
      warehouse.order_point,
      warehouseOrderPointDirection,
      orderPointStep,
      warehouseOrderPointDirection === "hold"
        ? `Warehouse order point already exceeds served-demand target; shortage risk points to timing or inbound pipeline, not a lower trigger point.`
        : `Warehouse coverage is below target or lost demand is present; ${inventoryReason}.`,
    );
    addNumericPolicy(
      "warehouse",
      "quantity",
      warehouse.quantity,
      warehouseQuantityDirection,
      quantityStep,
      warehouseQuantityDirection === "hold"
        ? `Keep warehouse quantity steady to avoid over-correction while inbound inventory catches up.`
        : "Increase outbound replenishment planning only after confirming inventory is available.",
    );

    if (
      String(factory.shipping_method_text || factory.shipping_method || "")
        .toLowerCase()
        .includes("truck")
    ) {
      addRecommendation({
        area: "factory",
        parameter: "shipping_method",
        baselineValue: policyMethodText(factory),
        suggestedValue: "emergency mail review only",
        direction: "review",
        urgency: Number.isFinite(lostDemand) && lostDemand > 0 ? "high" : "medium",
        confidence: "medium",
        reason: `Truck takes ${rules.truck_lead_days} days; consider temporary mail only for urgent recovery if inventory is below ${inventoryLow} and served lost demand persists.`,
      });
    }
  } else if (posture === "excess_stock") {
    addNumericPolicy(
      "factory",
      "order_point",
      factory.order_point,
      "decrease",
      orderPointStep,
      `Inventory coverage is above target; ${inventoryReason}.`,
    );
    addNumericPolicy(
      "factory",
      "quantity",
      factory.quantity,
      "decrease",
      quantityStep,
      "Reduce new inbound pressure to avoid holding excess stock.",
    );
    addNumericPolicy(
      "warehouse",
      "order_point",
      warehouse.order_point,
      "decrease",
      orderPointStep,
      `Warehouse has excess cover relative to the ${daysMin}-${daysMax} day target.`,
    );
    addNumericPolicy(
      "warehouse",
      "quantity",
      warehouse.quantity,
      "decrease",
      quantityStep,
      "Keep outbound settings conservative until stock cover normalizes.",
    );
  } else {
    addNumericPolicy(
      "factory",
      "order_point",
      factory.order_point,
      "hold",
      orderPointStep,
      `Coverage is within the configured target band; ${inventoryReason}.`,
    );
    addNumericPolicy(
      "factory",
      "quantity",
      factory.quantity,
      "hold",
      quantityStep,
      "No immediate production policy correction is indicated.",
    );
    addNumericPolicy(
      "warehouse",
      "order_point",
      warehouse.order_point,
      "hold",
      orderPointStep,
      "Warehouse policy appears stable against current demand and inventory.",
    );
  }

  if (shipmentBelowDemand && !shortageRisk) {
    addNumericPolicy(
      "warehouse",
      "quantity",
      warehouse.quantity,
      "increase",
      quantityStep,
      `Shipments trail demand; shipment/demand ratio is ${formatMetricNumber(shipmentRatio, "ratio")}.`,
    );
  }

  if (
    Number.isFinite(wipRatio) &&
    wipRatio > wipRatioMax &&
    excessCoverage &&
    !shortageRisk
  ) {
    addNumericPolicy(
      "factory",
      "quantity",
      factory.quantity,
      "decrease",
      quantityStep,
      `Factory WIP is high versus demand; WIP/demand ratio is ${formatMetricNumber(wipRatio, "ratio")}.`,
    );
  }

  const cashLeadRule = (config.monitor?.alert_rules || []).find(
    (rule) => rule.id === "cash_lead_narrow",
  );
  const cashLeadThreshold = Number(cashLeadRule?.threshold ?? 5);

  if (Number.isFinite(cashLead) && cashLead < cashLeadThreshold) {
    addRecommendation({
      area: "cash",
      parameter: "risk_posture",
      baselineValue: "normal",
      suggestedValue: "conservative",
      direction: "tighten",
      urgency: "medium",
      confidence: "medium",
      reason: `Cash lead over the nearest competitor is ${formatMetricNumber(cashLead, "%")}; prefer smaller policy moves until lead improves.`,
    });
  }

  return {
    enabled: true,
    mode,
    generated_at: record.checkedAt,
    generated_at_local: record.checkedAtLocal,
    target_team: record.targetTeam,
    dashboard_day: record.dashboardDay,
    posture,
    safety: {
      research_only: true,
      game_updates_enabled: false,
      submit_allowed: false,
      note: "This planner only writes recommendations to reports and files. It never submits Factory or Warehouse game forms.",
    },
    inputs: {
      warehouse_inventory: record.warehouseInventory,
      warehouse_inventory_day: record.warehouseDay,
      demand,
      shipments,
      lost_demand: lostDemand,
      lost_demand_rate: lostDemandRate,
      factory_wip: wip,
      days_of_cover: daysOfCover,
      shipment_to_demand_ratio: shipmentRatio,
      cash_lead_percent_vs_nearest: cashLead,
      target_inventory: targetInventory,
      target_rank: standingReport.target.rank,
      target_cash: standingReport.target.cash,
      baseline_source:
        scrapedBaseline.factory?.order_point !== undefined ||
        scrapedBaseline.warehouse?.order_point !== undefined
          ? "scraped policy pages"
          : "monitor_config fallback",
    },
    targets: {
      days_of_cover_min: daysMin,
      days_of_cover_max: daysMax,
      days_of_cover_target: daysTarget,
      warehouse_inventory_low: inventoryLow,
      warehouse_inventory_high: inventoryHigh,
      lost_demand_max: lostDemandMax,
      shipment_to_demand_ratio_min: shipmentRatioMin,
      wip_to_demand_ratio_max: wipRatioMax,
    },
    recommendations,
  };
}

function buildAdjustmentPlanCsv(plan) {
  const header = [
    "area",
    "parameter",
    "baseline",
    "suggested",
    "change",
    "direction",
    "urgency",
    "confidence",
    "reason",
    "submit_allowed",
  ];
  const rows = (plan?.recommendations || []).map((item) => [
    item.area,
    item.parameter,
    item.baseline,
    item.suggested,
    item.change,
    item.direction,
    item.urgency,
    item.confidence,
    item.reason,
    item.submit_allowed ? "yes" : "no",
  ]);

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
}

function policyApplyWorkflowUrl(config) {
  if (config.policy_apply?.workflow_url) {
    return config.policy_apply.workflow_url;
  }

  const repository = optionalEnv("GITHUB_REPOSITORY", "");
  return repository
    ? `https://github.com/${repository}/actions/workflows/apply-policy.yml`
    : "";
}

function controlNameForPolicyParameter(parameter) {
  return {
    shipping_method: "ship1",
    order_point: "point1",
    quantity: "quant1",
  }[parameter];
}

function pageIdForPolicyArea(area) {
  return {
    factory: "factory_calopeia",
    warehouse: "warehouse_calopeia",
  }[area];
}

function policyAreaCurrentValue(config, policySnapshot, area, parameter) {
  const baseline = policyBaseline(config, policySnapshot);
  return baseline?.[area]?.[parameter];
}

function strictShippingValue(value) {
  const normalized = String(value || "").trim().toLowerCase();

  if (normalized === "mail" || normalized === "truck") {
    return normalized;
  }

  return null;
}

function numericPolicyNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number) : null;
}

function samePolicyValue(parameter, left, right) {
  if (parameter === "shipping_method") {
    return strictShippingValue(left) === strictShippingValue(right);
  }

  return numericPolicyNumber(left) === numericPolicyNumber(right);
}

function candidatePolicyChangesFromPlan(config, plan, policySnapshot, options = {}) {
  const allowShipping = options.allowShipping === true;
  const candidates = new Map();
  const conflicts = [];

  for (const item of plan?.recommendations || []) {
    const area = String(item.area || "").toLowerCase();
    const parameter = String(item.parameter || "").toLowerCase();
    const control = controlNameForPolicyParameter(parameter);

    if (!["factory", "warehouse"].includes(area) || !control) {
      continue;
    }

    if (parameter === "shipping_method" && !allowShipping) {
      continue;
    }

    if (["hold", "review", "tighten"].includes(String(item.direction || "").toLowerCase())) {
      continue;
    }

    const current = policyAreaCurrentValue(config, policySnapshot, area, parameter);
    let suggested = item.suggested;

    if (parameter === "shipping_method") {
      suggested = strictShippingValue(suggested);
    } else {
      suggested = numericPolicyNumber(suggested);
    }

    if (suggested === null || suggested === undefined || samePolicyValue(parameter, current, suggested)) {
      continue;
    }

    const key = `${area}.${parameter}`;
    const candidate = {
      area,
      parameter,
      control,
      current,
      suggested,
      direction: item.direction || "",
      urgency: item.urgency || "",
      confidence: item.confidence || "",
      reason: item.reason || "",
      source: "latest recommendation",
    };

    if (candidates.has(key)) {
      const existing = candidates.get(key);
      if (!samePolicyValue(parameter, existing.suggested, candidate.suggested)) {
        conflicts.push({
          area,
          parameter,
          first: existing.suggested,
          second: candidate.suggested,
          reason: "Conflicting recommendations for the same field.",
        });
      }
      continue;
    }

    candidates.set(key, candidate);
  }

  return {
    changes: [...candidates.values()],
    conflicts,
  };
}

function policyWorkflowInputName(area, parameter) {
  const areaPrefix = String(area || "").toLowerCase();
  const field =
    parameter === "order_point"
      ? "order_point"
      : parameter === "quantity"
        ? "quantity"
        : parameter === "shipping_method"
          ? "shipping_method"
          : "";

  return areaPrefix && field ? `${areaPrefix}_${field}` : "";
}

function customPolicyChangesFromEnv(config, policySnapshot, baseSet = {}) {
  const specs = [
    ["POLICY_FACTORY_ORDER_POINT", "factory", "order_point"],
    ["POLICY_FACTORY_QUANTITY", "factory", "quantity"],
    ["POLICY_FACTORY_SHIPPING_METHOD", "factory", "shipping_method"],
    ["POLICY_WAREHOUSE_ORDER_POINT", "warehouse", "order_point"],
    ["POLICY_WAREHOUSE_QUANTITY", "warehouse", "quantity"],
    ["POLICY_WAREHOUSE_SHIPPING_METHOD", "warehouse", "shipping_method"],
  ];
  const changesByKey = new Map();
  const overriddenKeys = new Set();

  for (const change of baseSet.changes || []) {
    changesByKey.set(`${change.area}.${change.parameter}`, {
      ...change,
      source: change.source || "latest recommendation",
    });
  }

  for (const [envName, area, parameter] of specs) {
    const raw = optionalEnv(envName, "").trim();

    if (!raw) {
      continue;
    }

    const suggested =
      parameter === "shipping_method" ? strictShippingValue(raw) : numericPolicyNumber(raw);
    const current = policyAreaCurrentValue(config, policySnapshot, area, parameter);
    const key = `${area}.${parameter}`;
    const existing = changesByKey.get(key);
    overriddenKeys.add(key);

    changesByKey.set(key, {
      area,
      parameter,
      control: controlNameForPolicyParameter(parameter),
      current,
      suggested,
      direction:
        parameter === "shipping_method"
          ? "manual"
          : Number(suggested) > Number(current)
            ? "increase"
            : Number(suggested) < Number(current)
              ? "decrease"
              : "hold",
      urgency: "manual",
      confidence: "manual",
      reason: existing
        ? `Manual workflow input ${envName}; overrides latest suggested value ${existing.suggested}.`
        : `Manual workflow input ${envName}.`,
      source: existing ? "workflow override over latest recommendation" : "workflow input",
    });
  }

  return {
    changes: [...changesByKey.values()],
    conflicts: (baseSet.conflicts || []).filter(
      (conflict) => !overriddenKeys.has(`${conflict.area}.${conflict.parameter}`),
    ),
  };
}

function policyApplyCsv(report) {
  const header = [
    "area",
    "parameter",
    "current",
    "suggested",
    "delta",
    "source",
    "accepted",
    "applied",
    "verified",
    "reason",
  ];
  const rows = (report.changes || []).map((item) => [
    item.area,
    item.parameter,
    item.current ?? "",
    item.suggested ?? "",
    item.delta ?? "",
    item.source,
    item.accepted ? "yes" : "no",
    item.applied ? "yes" : "no",
    item.verified ? "yes" : "no",
    item.reject_reason || item.reason || "",
  ]);

  return `\uFEFF${[
    header.map(csvValue).join(","),
    ...rows.map((row) => row.map(csvValue).join(",")),
  ].join("\r\n")}\r\n`;
}

function policyApplyMarkdownSummary(report) {
  const changes = report.changes || [];
  const acceptedCount = changes.filter((item) => item.accepted).length;
  const appliedCount = changes.filter((item) => item.applied).length;
  const verifiedCount = changes.filter((item) => item.verified).length;
  const status = report.dry_run
    ? "DRY RUN ONLY - no game forms were submitted."
    : appliedCount > 0
      ? "GAME UPDATE SUBMITTED."
      : report.autopilot?.enabled
        ? "NO GAME UPDATE - autopilot guardrails blocked submit."
        : "NO GAME UPDATE - no accepted changes were submitted.";
  const rows = changes.map((item) =>
    [
      item.area,
      item.parameter,
      item.current ?? "",
      item.suggested ?? "",
      item.delta ?? "",
      item.accepted ? "yes" : "no",
      item.applied ? "yes" : "no",
      item.verified ? "yes" : "no",
      item.reject_reason || item.reason || "",
    ]
      .map((value) => String(value).replace(/\|/g, "\\|"))
      .join(" | "),
  );

  return [
    "## Policy Apply Result",
    "",
    `**Status:** ${status}`,
    "",
    `- Mode: ${report.mode}`,
    `- Confirm: ${report.confirm}`,
    `- Dry run: ${report.dry_run ? "yes" : "no"}`,
    report.autopilot?.enabled
      ? `- Autopilot: ${report.autopilot.apply_allowed ? "allowed" : "blocked"}; ${report.autopilot.reason}`
      : "",
    `- Target: ${report.target_team} rank ${report.target_rank ?? "n/a"} cash ${report.target_cash ?? "n/a"} day ${report.dashboard_day ?? "n/a"}`,
    `- Warehouse inventory: ${report.warehouse_inventory ?? "n/a"}`,
    `- Posture: ${report.posture ?? "n/a"}`,
    `- Conflicts: ${(report.conflicts || []).length}`,
    `- Accepted / applied / verified: ${acceptedCount} / ${appliedCount} / ${verifiedCount}`,
    "",
    "| Area | Field | Current | Suggested | Delta | Accepted | Applied | Verified | Reason |",
    "| --- | --- | ---: | ---: | ---: | --- | --- | --- | --- |",
    ...(rows.length ? rows.map((row) => `| ${row} |`) : ["| n/a | n/a |  |  |  | no | no | no | No candidate changes. |"]),
    "",
  ].join("\n");
}

function writePolicyApplyOutputs(config, report) {
  const jsonPath = path.resolve(
    process.cwd(),
    config.output.policy_apply_json || ".monitor-state/policy_apply_latest.json",
  );
  const csvPath = path.resolve(
    process.cwd(),
    config.output.policy_apply_csv || ".monitor-state/policy_apply_latest.csv",
  );

  ensureDir(path.dirname(jsonPath));
  fs.writeFileSync(jsonPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  fs.writeFileSync(csvPath, policyApplyCsv(report), "utf8");

  if (process.env.GITHUB_STEP_SUMMARY) {
    fs.appendFileSync(
      process.env.GITHUB_STEP_SUMMARY,
      `${policyApplyMarkdownSummary(report)}\n`,
      "utf8",
    );
  }

  return { jsonPath, csvPath };
}

function validatePolicyChanges(config, changes, context = {}) {
  const cfg = config.policy_apply || {};
  const guardrails = cfg.guardrails || {};
  const maxPerRun = Number(cfg.max_numeric_changes_per_run ?? 4);
  const maxChange = {
    ...(config.auto_adjust?.max_change_per_run || {}),
    ...(cfg.max_change_per_apply || {}),
  };
  const bounds = config.auto_adjust?.bounds || {};
  const metricCatalog = context.metricCatalog;
  const record = context.record;
  const coverTargets = coverageTargets(config, record);
  const rules = gameRules(config);
  const lostDemand =
    metricValue(metricCatalog, "derived:calopeia_served_lost_demand") ??
    metricValue(metricCatalog, "hq_lost_demand:Calopeia");
  const daysOfCover = metricValue(metricCatalog, "derived:days_of_cover");
  const allowShipping = context.allowShipping === true;
  let acceptedNumeric = 0;

  return changes.map((change) => {
    const result = {
      ...change,
      delta: "",
      accepted: false,
      applied: false,
      verified: false,
      reject_reason: "",
    };

    if (cfg.enabled === false) {
      result.reject_reason = "policy_apply.enabled is false.";
      return result;
    }

    if (!pageIdForPolicyArea(change.area) || !controlNameForPolicyParameter(change.parameter)) {
      result.reject_reason = "Unsupported policy field.";
      return result;
    }

    if (change.parameter === "priority" || guardrails.block_priority_changes !== false && /priority/i.test(change.parameter)) {
      result.reject_reason = "Priority level is not applied because the assignment states it has no effect.";
      return result;
    }

    if (change.parameter === "shipping_method") {
      const suggested = strictShippingValue(change.suggested);

      if (!suggested) {
        result.reject_reason = "Shipping method must be exactly mail or truck.";
        return result;
      }

      if (!allowShipping || cfg.allow_shipping_method_change !== true) {
        result.reject_reason = "Shipping method changes require explicit enablement and are disabled by default.";
        return result;
      }

      result.suggested = suggested;
      result.accepted = !samePolicyValue("shipping_method", change.current, suggested);
      result.reject_reason = result.accepted ? "" : "No change from current value.";
      return result;
    }

    const current = numericPolicyNumber(change.current);
    const suggested = numericPolicyNumber(change.suggested);

    if (!Number.isFinite(current) || !Number.isFinite(suggested)) {
      result.reject_reason = "Current and suggested values must be numeric.";
      return result;
    }

    const delta = suggested - current;
    result.current = current;
    result.suggested = suggested;
    result.delta = delta;

    if (delta === 0) {
      result.reject_reason = "No change from current value.";
      return result;
    }

    if (acceptedNumeric >= maxPerRun) {
      result.reject_reason = `Numeric change limit reached (${maxPerRun} per run).`;
      return result;
    }

    const fieldMaxChange = Number(maxChange[change.parameter] ?? 25);
    if (Math.abs(delta) > fieldMaxChange) {
      result.reject_reason = `Change ${delta} exceeds per-apply limit +/-${fieldMaxChange}.`;
      return result;
    }

    const min = Number(
      change.parameter === "quantity"
        ? bounds.quantity_min ?? 0
        : bounds.order_point_min ?? 0,
    );
    const max = Number(
      change.parameter === "quantity"
        ? bounds.quantity_max ?? 999999
        : bounds.order_point_max ?? 999999,
    );

    if (suggested < min || suggested > max) {
      result.reject_reason = `Suggested value ${suggested} is outside configured bounds ${min}-${max}.`;
      return result;
    }

    if (
      delta < 0 &&
      guardrails.block_decrease_when_lost_demand !== false &&
      Number.isFinite(lostDemand) &&
      lostDemand > 0
    ) {
      result.reject_reason = "Decrease blocked while served lost demand is present.";
      return result;
    }

    if (
      delta < 0 &&
      guardrails.block_decrease_below_dynamic_cover_min !== false &&
      Number.isFinite(daysOfCover) &&
      daysOfCover < coverTargets.min
    ) {
      result.reject_reason = "Decrease blocked because served cover is below the dynamic minimum.";
      return result;
    }

    if (
      delta > 0 &&
      guardrails.block_increase_above_dynamic_cover_max_without_lost_demand !== false &&
      Number.isFinite(daysOfCover) &&
      daysOfCover > coverTargets.max &&
      (!Number.isFinite(lostDemand) || lostDemand <= 0)
    ) {
      result.reject_reason = "Increase blocked because served cover is already above the dynamic maximum and no lost demand is present.";
      return result;
    }

    const currentShipping = policyAreaCurrentValue(
      config,
      context.policySnapshot || [],
      change.area,
      "shipping_method",
    );
    if (
      delta > 0 &&
      guardrails.block_truck_increase_inside_lead_time_window !== false &&
      strictShippingValue(currentShipping) === "truck" &&
      Number.isFinite(coverTargets.remainingDays) &&
      coverTargets.remainingDays <= rules.truck_lead_days
    ) {
      result.reject_reason = `Increase blocked because truck lead time is ${rules.truck_lead_days} days and only ${coverTargets.remainingDays} game days remain.`;
      return result;
    }

    acceptedNumeric += 1;
    result.accepted = true;
    return result;
  });
}

function policyChangeSignature(changes = []) {
  const parts = changes
    .filter((change) => change.accepted)
    .map((change) => {
      const delta = Number(change.delta);
      const direction = Number.isFinite(delta)
        ? delta > 0
          ? "increase"
          : delta < 0
            ? "decrease"
            : "hold"
        : String(change.direction || "").toLowerCase();
      return `${change.area}.${change.parameter}:${direction}`;
    })
    .filter(Boolean)
    .sort();

  return parts.join("|");
}

function buildAutopilotDecision(config, previousState, record, validated, candidateSet, options = {}) {
  const autopilot = config.policy_apply?.autopilot || {};
  const enabled = options.autopilot === true;
  const envName = autopilot.default_enabled_env || "POLICY_AUTOPILOT_ENABLED";
  const envText = optionalEnv(envName, "");
  const envEnabled = envText
    ? !["0", "false", "no", "off"].includes(envText.toLowerCase())
    : autopilot.enabled !== false;
  const accepted = validated.filter((change) => change.accepted);
  const currentDay = dashboardDayNumber(record);
  const lastApplyDay = dayNumberFromValue(previousState.last_autopilot_apply_day_number);
  const signature = policyChangeSignature(accepted);
  const previousSignature = previousState.last_autopilot_signature || "";
  const previousCount =
    previousSignature && previousSignature === signature
      ? Number(previousState.last_autopilot_signature_count || 0)
      : 0;
  const consecutiveCount = signature ? previousCount + 1 : 0;
  const requiredConfirmations = Math.max(
    1,
    Number(autopilot.consecutive_confirmations_required ?? 2) || 2,
  );
  const minGameDaysBetweenApply = Math.max(
    0,
    Number(autopilot.minimum_game_days_between_apply ?? 1) || 1,
  );
  const gameDayGap =
    Number.isFinite(currentDay) && Number.isFinite(lastApplyDay)
      ? currentDay - lastApplyDay
      : null;
  let applyAllowed = true;
  let reason = "Autopilot guardrails passed.";

  function block(message) {
    applyAllowed = false;
    reason = message;
  }

  if (!enabled) {
    block("Autopilot mode is not active for this run.");
  } else if (!envEnabled) {
    block(`${envName} disables autopilot.`);
  } else if (options.dryRun) {
    block("Dry run requested; no game forms will be submitted.");
  } else if (candidateSet.conflicts.length > 0) {
    block("Conflicting recommendations detected.");
  } else if (accepted.length === 0) {
    block("No accepted policy changes after guardrails.");
  } else if (!Number.isFinite(currentDay)) {
    block("Current dashboard day is unavailable.");
  } else if (
    Number.isFinite(lastApplyDay) &&
    gameDayGap !== null &&
    gameDayGap < minGameDaysBetweenApply
  ) {
    block(`Already applied on day ${lastApplyDay}; minimum spacing is ${minGameDaysBetweenApply} game day(s).`);
  } else if (!signature) {
    block("No stable policy-change signature.");
  } else if (consecutiveCount < requiredConfirmations) {
    block(`Need ${requiredConfirmations} consecutive matching recommendations; observed ${consecutiveCount}.`);
  } else if (remainingGameDays(config, record) <= 0) {
    block("Game has reached or passed the configured end day.");
  }

  return {
    enabled,
    env_name: envName,
    env_enabled: envEnabled,
    apply_allowed: applyAllowed,
    reason,
    current_signature: signature,
    previous_signature: previousSignature,
    consecutive_count: consecutiveCount,
    consecutive_required: requiredConfirmations,
    current_day: Number.isFinite(currentDay) ? currentDay : null,
    last_apply_day: Number.isFinite(lastApplyDay) ? lastApplyDay : null,
    minimum_game_days_between_apply: minGameDaysBetweenApply,
    accepted_change_count: accepted.length,
    stale_run_protection: "Every autopilot run re-crawls the latest game pages immediately before submit; workflow concurrency cancels older in-progress autopilot runs.",
  };
}

function policyPageConfig(config, area) {
  const pageId = pageIdForPolicyArea(area);
  const page = (config.crawl.policy_pages || []).find((item) => item.id === pageId);

  if (!page) {
    throw new Error(`Missing policy page configuration for ${area}`);
  }

  return page;
}

async function fetchPolicySnapshotForPage(config, cookieJar, page) {
  const response = await request(page.url, { method: "GET" }, cookieJar);
  const html = await response.text();

  if (!response.ok) {
    throw new Error(`${page.label} policy page failed with HTTP ${response.status}`);
  }

  return {
    html,
    snapshot: parsePolicyPageSnapshot(html, page),
  };
}

function formControlsForChanges(snapshot, changes) {
  const changeControls = new Set(changes.map((change) => change.control));
  const formIndex = (snapshot.forms || []).find((control) =>
    changeControls.has(control.name),
  )?.formIndex;

  if (!formIndex) {
    throw new Error(`Could not find an editable policy form for ${snapshot.id}`);
  }

  return (snapshot.forms || []).filter((control) => control.formIndex === formIndex);
}

function formBodyForPolicyChanges(controls, changes) {
  const updates = new Map(changes.map((change) => [change.control, String(change.suggested)]));
  const body = new URLSearchParams();

  for (const control of controls) {
    const name = control.name;

    if (!name) {
      continue;
    }

    const type = String(control.type || "").toLowerCase();

    if ((type === "checkbox" || type === "radio") && control.checked !== "yes") {
      continue;
    }

    if (type === "button") {
      continue;
    }

    const value = updates.has(name) ? updates.get(name) : control.value || "";
    body.append(name, value);
  }

  return body;
}

async function submitPolicyChangesForArea(config, cookieJar, area, changes) {
  const page = policyPageConfig(config, area);
  const before = await fetchPolicySnapshotForPage(config, cookieJar, page);
  const controls = formControlsForChanges(before.snapshot, changes);
  const body = formBodyForPolicyChanges(controls, changes);
  const firstControl = controls[0] || {};
  const method = String(firstControl.formMethod || "GET").toUpperCase();
  const actionUrl = new URL(firstControl.formAction || page.url, page.url);
  let response;

  if (method === "POST") {
    response = await follow(
      await request(
        actionUrl.toString(),
        {
          method: "POST",
          headers: {
            "content-type": "application/x-www-form-urlencoded",
          },
          body,
        },
        cookieJar,
      ),
      cookieJar,
    );
  } else {
    actionUrl.search = body.toString();
    response = await follow(
      await request(actionUrl.toString(), { method: "GET" }, cookieJar),
      cookieJar,
    );
  }

  const html = await response.text();

  if (!response.ok) {
    throw new Error(`${area} policy submit failed with HTTP ${response.status}`);
  }

  return {
    status: response.status,
    url: response.url,
    snapshot: parsePolicyPageSnapshot(html, page),
  };
}

function verifyAppliedChanges(config, policySnapshot, changes) {
  return changes.map((change) => {
    const current = policyAreaCurrentValue(
      config,
      policySnapshot,
      change.area,
      change.parameter,
    );

    return {
      ...change,
      verified: samePolicyValue(change.parameter, current, change.suggested),
      verified_value: current,
    };
  });
}

function addSummaryWorksheet(workbook, usedNames, config, record) {
  const worksheet = workbook.addWorksheet(safeWorksheetName("Summary", usedNames));
  const alpha = Number(config.excel?.exponential_smoothing_alpha ?? 0.3);

  worksheet.addRow(["Field", "Value"]);
  worksheet.addRow(["Generated at", record.checkedAtLocal]);
  worksheet.addRow(["Timezone", config.crawl.timezone]);
  worksheet.addRow(["Target team", record.targetTeam]);
  worksheet.addRow(["Target rank", record.targetRank]);
  worksheet.addRow(["Target cash", record.targetCash]);
  worksheet.addRow(["Dashboard day", record.dashboardDay]);
  worksheet.addRow(["Warehouse inventory", dataCell(record.warehouseInventory)]);
  worksheet.addRow(["Warehouse inventory day", dataCell(record.warehouseDay)]);
  worksheet.addRow(["Legacy inventory reference", dataCell(record.threshold)]);
  worksheet.addRow(["Inventory reference crossed", record.inventoryCheckpoint ? "yes" : "no"]);
  worksheet.addRow(["Active inventory alert", record.inventoryAlert ? "yes" : "no"]);
  worksheet.addRow(["EMA alpha", Number.isFinite(alpha) ? alpha : 0.3]);
  styleWorksheet(worksheet);
}

function addWatchlistWorksheet(workbook, usedNames, watchlist = []) {
  if (watchlist.length === 0) {
    return;
  }

  const worksheet = workbook.addWorksheet(safeWorksheetName("Watchlist", usedNames));
  worksheet.addRow([
    "rule_id",
    "indicator",
    "current_value",
    "unit",
    "operator",
    "threshold",
    "severity",
    "channels",
    "status",
    "message",
  ]);

  for (const item of watchlist) {
    const row = worksheet.addRow([
      item.id,
      item.label,
      Number.isFinite(item.current) ? item.current : "",
      item.unit,
      item.operator,
      Number.isFinite(item.threshold) ? item.threshold : "",
      item.severity,
      item.channels.join(", "),
      "",
      item.message,
    ]);
    const rowNumber = row.number;
    const statusFormula = [
      `IF(C${rowNumber}="","NO DATA",`,
      `IF(E${rowNumber}=">=",IF(C${rowNumber}>=F${rowNumber},"ALERT","OK"),`,
      `IF(E${rowNumber}=">",IF(C${rowNumber}>F${rowNumber},"ALERT","OK"),`,
      `IF(E${rowNumber}="<=",IF(C${rowNumber}<=F${rowNumber},"ALERT","OK"),`,
      `IF(E${rowNumber}="<",IF(C${rowNumber}<F${rowNumber},"ALERT","OK"),`,
      `IF(E${rowNumber}="=",IF(C${rowNumber}=F${rowNumber},"ALERT","OK"),"CHECK"))))))`,
    ].join("");
    row.getCell(9).value = {
      formula: statusFormula,
      result: item.status,
    };

    if (item.isAlert) {
      row.getCell(9).font = { bold: true, color: { argb: "FFB91C1C" } };
    }
  }

  styleWorksheet(worksheet);
}

function addAdjustmentPlanWorksheet(workbook, usedNames, plan) {
  const recommendations = plan?.recommendations || [];

  if (recommendations.length === 0) {
    return;
  }

  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Adjustment Plan", usedNames),
  );

  worksheet.addRow(["Field", "Value"]);
  worksheet.addRow(["Mode", plan.mode || "research_only"]);
  worksheet.addRow(["Posture", plan.posture || "n/a"]);
  worksheet.addRow(["Generated at", plan.generated_at_local || plan.generated_at || ""]);
  worksheet.addRow(["Game updates enabled", "no"]);
  worksheet.addRow(["Submit allowed", "no"]);
  worksheet.addRow(["Safety note", plan.safety?.note || "Research-only planner"]);
  worksheet.addRow([]);
  worksheet.addRow([
    "area",
    "parameter",
    "baseline",
    "suggested",
    "change",
    "direction",
    "urgency",
    "confidence",
    "reason",
    "submit_allowed",
  ]);

  for (const item of recommendations) {
    worksheet.addRow([
      item.area,
      item.parameter,
      item.baseline,
      item.suggested,
      item.change,
      item.direction,
      item.urgency,
      item.confidence,
      item.reason,
      item.submit_allowed ? "yes" : "no",
    ]);
  }

  styleWorksheet(worksheet);
  worksheet.getRow(9).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(9).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F2937" },
  };
}

function addPolicySummaryWorksheet(workbook, usedNames, policySnapshot = []) {
  if (!policySnapshot.length) {
    return;
  }

  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Policy Summary", usedNames),
  );
  worksheet.addRow([
    "page_id",
    "section",
    "region",
    "page_label",
    "item_type",
    "form_index",
    "name",
    "type",
    "value",
    "text",
    "checked",
  ]);

  for (const page of policySnapshot) {
    for (const fact of page.facts || []) {
      worksheet.addRow([
        page.id,
        page.section,
        page.region,
        page.label,
        "fact",
        "",
        "",
        "",
        fact,
        "",
        "",
      ]);
    }

    for (const control of page.forms || []) {
      worksheet.addRow([
        page.id,
        page.section,
        page.region,
        page.label,
        "form_control",
        control.formIndex,
        control.name,
        control.type,
        control.value,
        control.text,
        control.checked,
      ]);
    }
  }

  styleWorksheet(worksheet);
}

function addPolicyTablesWorksheet(workbook, usedNames, policySnapshot = []) {
  if (!policySnapshot.length) {
    return;
  }

  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Policy Tables", usedNames),
  );
  worksheet.addRow([
    "page_id",
    "section",
    "region",
    "page_label",
    "table_index",
    "table_label",
    "row_index",
    "column_index",
    "value",
  ]);

  for (const page of policySnapshot) {
    for (const table of page.tables || []) {
      for (const [rowIndex, row] of table.rows.entries()) {
        for (const [columnIndex, value] of row.entries()) {
          worksheet.addRow([
            page.id,
            page.section,
            page.region,
            page.label,
            table.index,
            table.label,
            rowIndex + 1,
            columnIndex + 1,
            value,
          ]);
        }
      }
    }
  }

  styleWorksheet(worksheet);
}

function addStandingWorksheet(workbook, usedNames, standingReport) {
  const worksheet = workbook.addWorksheet(
    safeWorksheetName("Team Standing", usedNames),
  );
  worksheet.addRow([
    "rank",
    "team",
    "cash",
    "cash_number",
    "gap_amount_target_minus_team",
    "gap_percent_vs_team_cash",
  ]);

  for (const row of standingReport.rows) {
    worksheet.addRow([
      dataCell(row.rank),
      row.team,
      row.cash,
      dataCell(row.cashNumber),
      row.gapAmountText,
      row.gapPercentText,
    ]);
  }

  styleWorksheet(worksheet);
}

function addPlotDataWorksheet(workbook, usedNames, plot, alpha = 0.3) {
  const worksheet = workbook.addWorksheet(
    safeWorksheetName(`${plot.section} ${plot.label}`, usedNames),
  );
  const originalColumnCount = plot.header.length;
  const seriesHeaders = plot.header.slice(1);
  const headers = [
    ...plot.header.map(dataCell),
    ...seriesHeaders.map((header) => `EMA ${header}`),
    ...seriesHeaders.map((header) => `Delta vs EMA ${header}`),
  ];
  const emaStartColumn = originalColumnCount + 1;
  const deltaStartColumn = emaStartColumn + seriesHeaders.length;
  const previousEma = Array(seriesHeaders.length).fill(null);

  worksheet.addRow(headers);

  for (const row of plot.rows) {
    const excelRow = worksheet.addRow([
      ...row.source.map(dataCell),
      ...Array(seriesHeaders.length * 2).fill(""),
    ]);
    const rowNumber = excelRow.number;

    for (let i = 0; i < seriesHeaders.length; i += 1) {
      const sourceColumn = i + 2;
      const emaColumn = emaStartColumn + i;
      const deltaColumn = deltaStartColumn + i;
      const sourceCell = `${columnLetter(sourceColumn)}${rowNumber}`;
      const emaCell = `${columnLetter(emaColumn)}${rowNumber}`;
      const previousEmaCell = `${columnLetter(emaColumn)}${rowNumber - 1}`;
      const rawValue = row.raw[sourceColumn - 1];
      const hasValue = Number.isFinite(rawValue);
      const currentPrevious = previousEma[i];
      const emaResult = hasValue
        ? currentPrevious === null
          ? rawValue
          : alpha * rawValue + (1 - alpha) * currentPrevious
        : currentPrevious;
      const emaFormula =
        rowNumber === 2
          ? `IF(${sourceCell}="","",VALUE(${sourceCell}))`
          : `IF(${sourceCell}="",${previousEmaCell},'Summary'!$B$12*VALUE(${sourceCell})+(1-'Summary'!$B$12)*${previousEmaCell})`;
      const deltaFormula = `IF(OR(${sourceCell}="",${emaCell}=""),"",VALUE(${sourceCell})-${emaCell})`;

      excelRow.getCell(emaColumn).value = {
        formula: emaFormula,
        result: Number.isFinite(emaResult) ? emaResult : "",
      };

      excelRow.getCell(deltaColumn).value = {
        formula: deltaFormula,
        result:
          hasValue && Number.isFinite(emaResult) ? rawValue - emaResult : "",
      };

      if (Number.isFinite(emaResult)) {
        previousEma[i] = emaResult;
      }
    }
  }

  styleWorksheet(worksheet);
}

function dataWorkbookFilename(record, options = {}) {
  const prefix = options.test ? "test-" : "";
  const day = safeAttachmentFilename(record.dashboardDay);
  const team = safeAttachmentFilename(record.targetTeam);

  return `${prefix}mgt267-data-${team}-day-${day}.xlsx`;
}

async function buildDataWorkbookBuffer(
  config,
  record,
  standingReport,
  plotSnapshots,
  options = {},
) {
  if (!plotSnapshots || plotSnapshots.length === 0) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "MGT267 Watchdog";
  workbook.created = new Date(record.checkedAt);
  workbook.modified = new Date(record.checkedAt);
  workbook.calcProperties.fullCalcOnLoad = true;
  const usedNames = new Set();
  const alpha = Number(config.excel?.exponential_smoothing_alpha ?? 0.3);
  const operationalSnapshot =
    options.operationalSnapshot || buildOperationalSnapshot(plotSnapshots, {});
  const metricCatalog =
    options.metricCatalog ||
    buildMetricCatalog(config, record, standingReport, operationalSnapshot, {
      policySnapshot: options.policySnapshot || [],
    });
  const watchlist =
    options.watchlist || buildWatchlist(config, metricCatalog, "hourly");
  const adjustmentPlan =
    options.adjustmentPlan ||
    buildAutoAdjustmentPlan(
      config,
      metricCatalog,
      record,
      standingReport,
      options.policySnapshot,
    );
  const backtestReport =
    options.backtestReport ||
    buildBacktestReport(
      config,
      record,
      standingReport,
      plotSnapshots,
      options.policySnapshot || [],
    );

  addSummaryWorksheet(workbook, usedNames, config, record);
  addWatchlistWorksheet(workbook, usedNames, watchlist);
  addAdjustmentPlanWorksheet(workbook, usedNames, adjustmentPlan);
  addBacktestWorksheets(workbook, usedNames, backtestReport);
  addPolicySummaryWorksheet(workbook, usedNames, options.policySnapshot || []);
  addPolicyTablesWorksheet(workbook, usedNames, options.policySnapshot || []);

  for (const plot of plotSnapshots) {
    addPlotDataWorksheet(workbook, usedNames, plot, Number.isFinite(alpha) ? alpha : 0.3);
  }

  addStandingWorksheet(workbook, usedNames, standingReport);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
}

async function writeDataWorkbookFile(
  config,
  record,
  standingReport,
  plotSnapshots,
  filePath,
  options = {},
) {
  const buffer = await buildDataWorkbookBuffer(
    config,
    record,
    standingReport,
    plotSnapshots,
    options,
  );

  if (!buffer) {
    return false;
  }

  fs.writeFileSync(filePath, buffer);
  return true;
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
    "inventory_checkpoint",
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
    record.inventoryCheckpoint ? "yes" : "no",
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

function isTruthyEnv(name) {
  return ["1", "true", "yes"].includes(optionalEnv(name, "false").toLowerCase());
}

function shouldSendReportNow(config, previousState) {
  if (isTruthyEnv("FORCE_EMAIL_REPORT")) {
    return {
      send: true,
      reason: "forced by workflow trigger",
    };
  }

  const intervalMinutes = Number(config.monitor.report_min_interval_minutes || 0);

  if (!Number.isFinite(intervalMinutes) || intervalMinutes <= 0) {
    return {
      send: true,
      reason: "no report interval throttle configured",
    };
  }

  if (!previousState.last_email_sent_at) {
    return {
      send: true,
      reason: "no previous email timestamp",
    };
  }

  const lastSentAt = new Date(previousState.last_email_sent_at).getTime();

  if (!Number.isFinite(lastSentAt)) {
    return {
      send: true,
      reason: "invalid previous email timestamp",
    };
  }

  const elapsedMinutes = (Date.now() - lastSentAt) / 60000;

  return {
    send: elapsedMinutes >= intervalMinutes,
    reason: `last email ${elapsedMinutes.toFixed(1)} minutes ago; minimum interval ${intervalMinutes} minutes`,
  };
}

function normalizeRecipients(recipients) {
  return [...new Set((recipients || []).map((item) => String(item).trim()).filter(Boolean))];
}

function formatFromAddress() {
  const from = requiredEnv("SMTP_FROM");

  if (from.includes("<")) {
    return from;
  }

  return `MGT267 Watchdog <${from}>`;
}

function emailDeliveryLogPath(config) {
  return path.resolve(
    process.cwd(),
    config.output.email_delivery_json || ".monitor-state/email_delivery_latest.json",
  );
}

function writeEmailDeliveryLog(config, deliveryLog) {
  const filePath = emailDeliveryLogPath(config);
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(`${filePath}`, `${JSON.stringify(deliveryLog, null, 2)}\n`, "utf8");
  return filePath;
}

function markdownCell(value) {
  return String(value ?? "")
    .replace(/\r?\n/g, " ")
    .replace(/\|/g, "\\|")
    .trim();
}

function actionsRunUrl() {
  const serverUrl = optionalEnv("GITHUB_SERVER_URL", "https://github.com");
  const repository = optionalEnv("GITHUB_REPOSITORY", "");
  const runId = optionalEnv("GITHUB_RUN_ID", "");

  return repository && runId ? `${serverUrl}/${repository}/actions/runs/${runId}` : "";
}

function githubRepositoryUrl() {
  const serverUrl = optionalEnv("GITHUB_SERVER_URL", "https://github.com");
  const repository = optionalEnv("GITHUB_REPOSITORY", "SammyFang/MGT267watchdog");

  return repository ? `${serverUrl}/${repository}` : "";
}

function statusPageUrl() {
  return optionalEnv("STATUS_PAGE_URL", "");
}

function dayNumberFromValue(value) {
  const text = String(value ?? "").replace(/,/g, "").trim();

  if (!text) {
    return null;
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

function gamePhaseLabel(config, record) {
  const rules = gameRules(config);
  const remaining = remainingGameDays(config, record);

  if (!Number.isFinite(remaining)) {
    return "unknown";
  }

  if (remaining <= 0) {
    return "ended";
  }

  if (remaining <= Number(rules.truck_lead_days || 7)) {
    return "final lead-time window";
  }

  if (remaining <= 30) {
    return "final month";
  }

  if (remaining <= Number(rules.capacity_expansion_lead_days || 90)) {
    return "endgame capacity cutoff";
  }

  if (remaining <= Number(rules.capacity_expansion_lead_days || 90) + 30) {
    return "late game";
  }

  return "normal operations";
}

function buildDayChangeReview(config, previousState, record, metricCatalog, adjustmentPlan, backtestReport) {
  const currentDay = dashboardDayNumber(record);
  const previousDay = dayNumberFromValue(
    previousState.last_dashboard_day ?? previousState.last_day_change_dashboard_day,
  );
  const changed =
    Number.isFinite(currentDay) &&
    (!Number.isFinite(previousDay) || currentDay !== previousDay);
  const advancedBy =
    Number.isFinite(currentDay) && Number.isFinite(previousDay)
      ? currentDay - previousDay
      : null;
  const rules = gameRules(config);
  const remainingDays = remainingGameDays(config, record);
  const targets = coverageTargets(config, record);
  const latest = backtestReport?.latest || {};
  const daysOfCover =
    metricRaw(metricCatalog, "derived:days_of_cover") ||
    ratioMetric(latest.days_of_cover) ||
    "n/a";
  const lostDemand =
    metricRaw(metricCatalog, "derived:calopeia_served_lost_demand") ||
    String(latest.lost_demand ?? "n/a");
  const demand =
    metricRaw(metricCatalog, "derived:calopeia_served_demand") ||
    String(latest.demand ?? "n/a");
  const shipments =
    metricRaw(metricCatalog, "derived:calopeia_served_shipments") ||
    String(latest.shipments ?? "n/a");
  const shipmentRatio =
    metricRaw(metricCatalog, "derived:shipment_to_demand_ratio") ||
    ratioMetric(latest.shipment_to_demand_ratio) ||
    "n/a";
  const cashLead =
    metricRaw(metricCatalog, "derived:cash_lead_percent_vs_nearest") || "n/a";
  const maxChange = config.policy_apply?.max_change_per_apply ||
    config.auto_adjust?.max_change_per_run ||
    {};
  const candidateActions = (adjustmentPlan?.recommendations || [])
    .filter((item) => !["hold", "review"].includes(String(item.direction || "").toLowerCase()))
    .slice(0, 4)
    .map((item) => `${item.area}.${item.parameter}: ${item.baseline || "n/a"} -> ${item.suggested || "n/a"} (${item.direction || "n/a"})`);
  const backtestCandidates = (backtestReport?.recommendations || [])
    .slice(0, 4)
    .map((item) =>
      `${item.indicator}: ${item.operator} ${item.suggested_threshold}; precision=${
        Number.isFinite(item.precision) ? percentMetric(item.precision * 100) : "n/a"
      }, recall=${Number.isFinite(item.recall) ? percentMetric(item.recall * 100) : "n/a"}`,
    );
  const guardrails = [
    `Max policy move per apply: order_point +/-${maxChange.order_point ?? 25}, quantity +/-${maxChange.quantity ?? 25}; avoid reacting to one noisy day.`,
    `Do not lower order point/quantity when served lost demand is above 0 or cover is below ${ratioMetric(targets.min)} days.`,
    `Do not raise production aggressively when cover is above ${ratioMetric(targets.max)} days and lost demand is 0; trim or increase gradually to avoid bullwhip.`,
    `Capacity takes ${rules.capacity_expansion_lead_days} days and cannot be retired; avoid new capacity in ${gamePhaseLabel(config, record)} unless backtest shows persistent lost demand.`,
    `Game ends on day ${rules.end_day}; remaining inventory and capacity are obsolete after that date.`,
  ];

  return {
    enabled: true,
    changed,
    previous_day: Number.isFinite(previousDay) ? previousDay : null,
    current_day: Number.isFinite(currentDay) ? currentDay : null,
    advanced_by: Number.isFinite(advancedBy) ? advancedBy : null,
    remaining_days: remainingDays,
    end_day: rules.end_day,
    phase: gamePhaseLabel(config, record),
    posture: adjustmentPlan?.posture || "n/a",
    latest: {
      warehouse_inventory: record.warehouseInventory,
      demand,
      shipments,
      lost_demand: lostDemand,
      days_of_cover: daysOfCover,
      shipment_to_demand_ratio: shipmentRatio,
      cash_lead_percent_vs_nearest: cashLead,
    },
    backtest: {
      observations: backtestReport?.observations ?? 0,
      window: backtestReport?.day_start && backtestReport?.day_end
        ? `${backtestReport.day_start} to ${backtestReport.day_end}`
        : "n/a",
      horizon_days: backtestReport?.horizon_days ?? "n/a",
      recommendations: backtestCandidates,
    },
    action_candidates: candidateActions,
    guardrails,
  };
}

function buildDayChangeReviewMarkdown(review) {
  if (!review?.enabled) {
    return "";
  }

  const changeText = review.changed
    ? `changed from ${review.previous_day ?? "n/a"} to ${review.current_day ?? "n/a"}`
    : `unchanged at ${review.current_day ?? "n/a"}`;

  return [
    "### Simulated Day Review",
    "",
    `- Day status: ${changeText}${Number.isFinite(review.advanced_by) ? ` (${review.advanced_by >= 0 ? "+" : ""}${review.advanced_by})` : ""}`,
    `- Game clock: day ${review.current_day ?? "n/a"} / ${review.end_day}; remaining days ${review.remaining_days ?? "n/a"}; phase ${review.phase}`,
    `- Posture: ${review.posture}`,
    `- Latest operations: inventory ${review.latest.warehouse_inventory ?? "n/a"}, demand ${review.latest.demand}, shipments ${review.latest.shipments}, lost demand ${review.latest.lost_demand}, cover ${review.latest.days_of_cover}, shipment/demand ${review.latest.shipment_to_demand_ratio}`,
    `- Cash lead vs nearest: ${review.latest.cash_lead_percent_vs_nearest}`,
    "",
    review.action_candidates.length
      ? `Action candidates: ${review.action_candidates.join("; ")}`
      : "Action candidates: none; hold current policy unless alerts worsen.",
    review.backtest.recommendations.length
      ? `Backtest thresholds: ${review.backtest.recommendations.join("; ")}`
      : `Backtest thresholds: no strong historical threshold in ${review.backtest.window}.`,
    `Backtest window: ${review.backtest.window}; observations ${review.backtest.observations}; horizon ${review.backtest.horizon_days} days.`,
    "",
    "Bullwhip / endgame guardrails:",
    ...review.guardrails.map((item) => `- ${item}`),
    "",
  ].join("\n");
}

function buildDayChangeReviewLines(review) {
  const markdown = buildDayChangeReviewMarkdown(review).trim();
  return markdown ? ["", ...markdown.split(/\n/)] : [];
}

function buildWatchdogMarkdownSummary(config, record, standingReport, options = {}) {
  const kind = options.kind === "warning" ? "15-minute warning check" : "hourly monitor";
  const watchlist = options.watchlist || [];
  const alerts = options.metricAlerts || watchlist.filter((item) => item.isAlert);
  const criticalCount = alerts.filter((item) => item.severity === "critical").length;
  const runUrl = actionsRunUrl();
  const gameUrl = config.email?.game_entry_url || config.crawl.entry_url;
  const adjustmentItems = (options.adjustmentPlan?.recommendations || []).slice(0, 4);
  const standingRows = (standingReport.rows || []).slice(0, 8);
  const status =
    alerts.length > 0
      ? `${alerts.length} alert(s), ${criticalCount} critical`
      : "No active warning alerts";

  return [
    `## MGT267 Watchdog ${kind}`,
    "",
    `**Status:** ${status}`,
    "",
    `- Checked at: ${record.checkedAtLocal || record.checkedAt || "n/a"} ${config.crawl.timezone || ""}`.trim(),
    `- Team: ${record.targetTeam || "n/a"}`,
    `- Rank: ${record.targetRank ?? "n/a"}`,
    `- Cash: ${record.targetCash || "n/a"}`,
    `- Dashboard day: ${record.dashboardDay || "n/a"}`,
    `- Warehouse inventory: ${record.warehouseInventory ?? "n/a"}`,
    `- Email sent: ${options.emailSent ? "yes" : "no"}${isEmailEnabled(config) ? "" : " (SMTP disabled for Gmail safety)"}`,
    options.emailError ? `- Email error: ${options.emailError}` : "",
    gameUrl ? `- Game entry: ${gameUrl}` : "",
    runUrl ? `- GitHub run: ${runUrl}` : "",
    "",
    buildDayChangeReviewMarkdown(options.dayChangeReview),
    "### Alerts",
    alerts.length
      ? "| Severity | Alert | Current | Rule | Message |\n| --- | --- | ---: | --- | --- |\n" +
        alerts
          .map((alert) =>
            `| ${markdownCell(alert.severity)} | ${markdownCell(alert.label)} | ${markdownCell(alert.currentRaw || alert.current)} | ${markdownCell(`${alert.operator} ${alert.thresholdRaw}`)} | ${markdownCell(alert.message)} |`,
          )
          .join("\n")
      : "No active alerts on this check.",
    "",
    "### Recommended Policy Changes",
    adjustmentItems.length
      ? "| Area | Field | Baseline | Suggested | Direction | Reason |\n| --- | --- | ---: | ---: | --- | --- |\n" +
        adjustmentItems
          .map((item) =>
            `| ${markdownCell(item.area)} | ${markdownCell(item.parameter)} | ${markdownCell(item.baseline)} | ${markdownCell(item.suggested)} | ${markdownCell(item.direction)} | ${markdownCell(item.reason)} |`,
          )
          .join("\n")
      : "No policy changes recommended.",
    "",
    "### Standing Snapshot",
    standingRows.length
      ? "| Rank | Team | Cash | Gap | Gap % |\n| ---: | --- | ---: | ---: | ---: |\n" +
        standingRows
          .map((row) =>
            `| ${markdownCell(row.rank)} | ${markdownCell(row.team)} | ${markdownCell(row.cash)} | ${markdownCell(row.gapAmountText)} | ${markdownCell(row.gapPercentText)} |`,
          )
          .join("\n")
      : "No standing rows available.",
    "",
    "Artifacts contain the latest workbook, policy snapshot, backtest, and raw crawl outputs when generated.",
    "",
  ]
    .join("\n");
}

function statusSiteDir(config) {
  return path.resolve(
    process.cwd(),
    config.output.status_site_dir || ".monitor-state/status-site",
  );
}

function statusPageTable(headers, rows) {
  if (!rows.length) {
    return '<div class="empty">No data available.</div>';
  }

  return [
    "<table>",
    "<thead><tr>",
    ...headers.map((header) => `<th>${escapeHtml(header)}</th>`),
    "</tr></thead><tbody>",
    ...rows.map((row) =>
      `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`,
    ),
    "</tbody></table>",
  ].join("");
}

function compactAlertRows(alerts = []) {
  return alerts.map((alert) => [
    alert.severity || "",
    alert.label || "",
    alert.currentRaw || alert.current || "",
    `${alert.operator || ""} ${alert.thresholdRaw || alert.threshold || ""}`.trim(),
    alert.message || "",
  ]);
}

function compactAdjustmentRows(adjustmentPlan = {}) {
  return (adjustmentPlan.recommendations || []).slice(0, 8).map((item) => [
    item.area || "",
    item.parameter || "",
    item.baseline || "",
    item.suggested || "",
    item.direction || "",
    item.reason || "",
  ]);
}

function compactStandingRows(standingReport = {}) {
  return (standingReport.rows || []).slice(0, 12).map((row) => [
    row.rank || "",
    row.team || "",
    row.cash || "",
    row.gapAmountText || "",
    row.gapPercentText || "",
  ]);
}

function dayChangeRows(review = {}) {
  if (!review.enabled) {
    return [];
  }

  return [
    ["Day changed", review.changed ? "yes" : "no"],
    ["Previous day", review.previous_day ?? "n/a"],
    ["Current day", review.current_day ?? "n/a"],
    ["Advanced by", Number.isFinite(review.advanced_by) ? review.advanced_by : "n/a"],
    ["Remaining days", review.remaining_days ?? "n/a"],
    ["Phase", review.phase || "n/a"],
    ["Posture", review.posture || "n/a"],
    ["Inventory", review.latest?.warehouse_inventory ?? "n/a"],
    ["Demand / shipments", `${review.latest?.demand ?? "n/a"} / ${review.latest?.shipments ?? "n/a"}`],
    ["Lost demand", review.latest?.lost_demand ?? "n/a"],
    ["Days of cover", review.latest?.days_of_cover ?? "n/a"],
    ["Backtest window", review.backtest?.window || "n/a"],
  ];
}

function buildStatusPageHtml(config, record, standingReport, options = {}) {
  const alerts = options.metricAlerts || [];
  const alertCount = alerts.length;
  const criticalCount = alerts.filter((item) => item.severity === "critical").length;
  const statusClass = criticalCount > 0 ? "critical" : alertCount > 0 ? "warning" : "ok";
  const statusText =
    alertCount > 0
      ? `${alertCount} active alert${alertCount === 1 ? "" : "s"} (${criticalCount} critical)`
      : "No active alerts";
  const issueUrl = options.githubIssueUrl || "";
  const pageUrl = statusPageUrl();
  const runUrl = actionsRunUrl();
  const repoUrl = githubRepositoryUrl();
  const gameUrl = config.email?.game_entry_url || config.crawl.entry_url;
  const applyUrl = policyApplyWorkflowUrl(config);
  const generatedAt = record.checkedAtLocal || record.checkedAt || "n/a";

  return [
    "<!doctype html>",
    '<html lang="en">',
    "<head>",
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    '<meta http-equiv="refresh" content="60">',
    "<title>MGT267 Watchdog Live Status</title>",
    "<style>",
    ":root{color-scheme:light;--bg:#f6f7f9;--ink:#111827;--muted:#5b6472;--line:#d7dde6;--card:#fff;--blue:#1d4ed8;--green:#047857;--red:#b91c1c;--amber:#b45309}",
    "*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Arial,Helvetica,sans-serif}main{max-width:1160px;margin:0 auto;padding:22px}",
    ".top{display:flex;gap:16px;align-items:flex-start;justify-content:space-between;margin-bottom:16px}.title h1{font-size:26px;line-height:1.1;margin:0}.title p{margin:6px 0 0;color:var(--muted)}",
    ".pill{display:inline-flex;align-items:center;border-radius:999px;padding:8px 12px;font-size:13px;font-weight:700;border:1px solid var(--line);background:#fff}.pill.ok{color:var(--green)}.pill.warning{color:var(--amber)}.pill.critical{color:var(--red)}",
    ".grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:14px 0}.metric{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px}.metric .label{font-size:12px;color:var(--muted);text-transform:uppercase}.metric .value{font-size:22px;font-weight:700;margin-top:4px;overflow-wrap:anywhere}",
    ".links{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 18px}.links a{display:inline-block;text-decoration:none;color:#fff;background:var(--blue);border-radius:6px;padding:9px 12px;font-size:13px;font-weight:700}.links a.secondary{background:#334155}",
    "section{background:var(--card);border:1px solid var(--line);border-radius:8px;margin:12px 0;overflow:hidden}section h2{font-size:17px;margin:0;padding:12px 14px;border-bottom:1px solid var(--line);background:#f8fafc}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:9px 10px;border-bottom:1px solid #e7ebf1;text-align:left;vertical-align:top}th{font-size:12px;color:#334155;background:#f8fafc}td{overflow-wrap:anywhere}.empty{padding:12px 14px;color:var(--muted)}",
    ".downloads a{color:var(--blue);font-weight:700}.foot{color:var(--muted);font-size:12px;margin:18px 0 4px}",
    "@media(max-width:760px){main{padding:14px}.top{display:block}.grid{grid-template-columns:repeat(2,minmax(0,1fr))}table{font-size:12px}}",
    "</style>",
    "</head>",
    "<body><main>",
    '<div class="top">',
    '<div class="title">',
    "<h1>MGT267 Watchdog Live Status</h1>",
    `<p>Auto-refreshes every 60 seconds. Last updated ${escapeHtml(generatedAt)} ${escapeHtml(config.crawl.timezone || "")}</p>`,
    "</div>",
    `<div class="pill ${statusClass}">${escapeHtml(statusText)}</div>`,
    "</div>",
    '<div class="grid">',
    `<div class="metric"><div class="label">Team</div><div class="value">${escapeHtml(record.targetTeam || "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Rank</div><div class="value">${escapeHtml(record.targetRank ?? "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Cash</div><div class="value">${escapeHtml(record.targetCash || "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Day</div><div class="value">${escapeHtml(record.dashboardDay || "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Warehouse Inventory</div><div class="value">${escapeHtml(record.warehouseInventory ?? "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Inventory Day</div><div class="value">${escapeHtml(record.warehouseDay || "n/a")}</div></div>`,
    `<div class="metric"><div class="label">Email</div><div class="value">${escapeHtml(options.emailSent ? "sent" : "off")}</div></div>`,
    `<div class="metric"><div class="label">Run Type</div><div class="value">${escapeHtml(options.kind === "warning" ? "15m" : "hourly")}</div></div>`,
    "</div>",
    '<div class="links">',
    gameUrl ? `<a href="${escapeHtml(gameUrl)}">Open Game</a>` : "",
    pageUrl ? `<a href="${escapeHtml(pageUrl)}" class="secondary">Refresh Page</a>` : "",
    applyUrl ? `<a href="${escapeHtml(applyUrl)}" class="secondary">Apply Workflow</a>` : "",
    issueUrl ? `<a href="${escapeHtml(issueUrl)}" class="secondary">Live Issue</a>` : "",
    runUrl ? `<a href="${escapeHtml(runUrl)}" class="secondary">Current Run</a>` : "",
    repoUrl ? `<a href="${escapeHtml(repoUrl)}/actions" class="secondary">Actions</a>` : "",
    "</div>",
    "<section><h2>Simulated Day Review</h2>",
    statusPageTable(["Metric", "Value"], dayChangeRows(options.dayChangeReview)),
    "</section>",
    "<section><h2>Alerts</h2>",
    statusPageTable(["Severity", "Alert", "Current", "Rule", "Message"], compactAlertRows(alerts)),
    "</section>",
    "<section><h2>Recommended Policy Changes</h2>",
    statusPageTable(
      ["Area", "Field", "Current", "Suggested", "Direction", "Reason"],
      compactAdjustmentRows(options.adjustmentPlan),
    ),
    "</section>",
    "<section><h2>Standing Snapshot</h2>",
    statusPageTable(["Rank", "Team", "Cash", "Gap", "Gap %"], compactStandingRows(standingReport)),
    "</section>",
    '<section class="downloads"><h2>Downloads</h2><div class="empty">',
    '<a href="status.json">status.json</a> | <a href="summary.md">summary.md</a> | <a href="supply_chain_data_latest.xlsx">data workbook</a> | <a href="backtest_report_latest.xlsx">backtest workbook</a>',
    "</div></section>",
    '<div class="foot">Developed by Yung-Sian Fang. Gmail SMTP is intentionally disabled unless EMAIL_ENABLED is set true.</div>',
    "</main></body></html>",
  ].join("");
}

function statusJsonPayload(record, standingReport, options = {}) {
  return {
    generated_at: record.checkedAt,
    generated_at_local: record.checkedAtLocal,
    kind: options.kind || "hourly",
    target_team: record.targetTeam,
    target_rank: record.targetRank,
    target_cash: record.targetCash,
    dashboard_day: record.dashboardDay,
    warehouse_inventory: record.warehouseInventory,
    warehouse_inventory_day: record.warehouseDay,
    email_sent: Boolean(options.emailSent),
    email_error: options.emailError || "",
    day_change_review: options.dayChangeReview || null,
    alerts: options.metricAlerts || [],
    adjustment_plan: options.adjustmentPlan || null,
    standing: standingReport.rows || [],
    action_run_url: actionsRunUrl(),
    game_entry_url: options.config?.email?.game_entry_url || "",
  };
}

function copyIfExists(sourcePath, destinationPath) {
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    return false;
  }

  ensureDir(path.dirname(destinationPath));
  fs.copyFileSync(sourcePath, destinationPath);
  return true;
}

function writeStatusSite(config, record, standingReport, options = {}, markdown = "") {
  const siteDir = statusSiteDir(config);
  ensureDir(siteDir);
  const html = buildStatusPageHtml(config, record, standingReport, options);
  const payload = statusJsonPayload(record, standingReport, {
    ...options,
    config,
  });

  fs.writeFileSync(path.join(siteDir, "index.html"), html, "utf8");
  fs.writeFileSync(path.join(siteDir, "summary.md"), `${markdown}\n`, "utf8");
  fs.writeFileSync(path.join(siteDir, "status.json"), `${JSON.stringify(payload, null, 2)}\n`, "utf8");

  const downloadable = [
    [config.output.data_workbook_xlsx, "supply_chain_data_latest.xlsx"],
    [config.output.backtest_xlsx, "backtest_report_latest.xlsx"],
    [config.output.adjustment_plan_csv, "adjustment_plan_latest.csv"],
    [config.output.policy_snapshot_csv, "policy_snapshot_latest.csv"],
    [config.output.standing_gaps_csv, "standing_gaps_latest.csv"],
  ];

  for (const [source, filename] of downloadable) {
    copyIfExists(path.resolve(process.cwd(), source || ""), path.join(siteDir, filename));
  }

  return siteDir;
}

function truncateText(text, maxLength) {
  const value = String(text || "");
  return value.length > maxLength ? `${value.slice(0, maxLength - 3)}...` : value;
}

function shouldSendWecom(config, options = {}) {
  const cfg = config.notifications?.wecom || {};

  if (cfg.enabled === false || !isTruthyEnv("WECOM_ENABLED")) {
    return false;
  }

  const mode = optionalEnv("WECOM_NOTIFY_MODE", cfg.notify_mode || "alerts").toLowerCase();
  const alerts = options.metricAlerts || [];
  const dayChanged = options.dayChangeReview?.changed === true;

  return (
    mode === "all" ||
    (mode === "day_change" && dayChanged) ||
    (mode === "alerts_or_day_change" && (alerts.length > 0 || dayChanged)) ||
    (mode === "alerts" && alerts.length > 0)
  );
}

function buildWecomMarkdown(config, record, standingReport, options = {}) {
  const alerts = options.metricAlerts || [];
  const alertLine = alerts.length
    ? alerts.map((alert) => `> ${alert.severity || "alert"} ${alert.label}: ${alert.currentRaw || alert.current || "n/a"} ${alert.operator || ""} ${alert.thresholdRaw || ""}`).join("\n")
    : "> No active alerts";
  const topRecommendation = (options.adjustmentPlan?.recommendations || [])[0];
  const review = options.dayChangeReview || {};
  const pageUrl = statusPageUrl();
  const issueUrl = options.githubIssueUrl || "";
  const runUrl = actionsRunUrl();

  return truncateText(
    [
      `**MGT267 Watchdog ${options.kind === "warning" ? "15m" : "hourly"}**`,
      `Team ${record.targetTeam || "n/a"} rank ${record.targetRank ?? "n/a"} cash ${record.targetCash || "n/a"} day ${record.dashboardDay || "n/a"}`,
      `Warehouse inventory: ${record.warehouseInventory ?? "n/a"}`,
      review.enabled
        ? `Sim day review: ${review.changed ? "changed" : "unchanged"} ${review.previous_day ?? "n/a"} -> ${review.current_day ?? "n/a"}; remaining ${review.remaining_days ?? "n/a"}; phase ${review.phase || "n/a"}`
        : "",
      "",
      alertLine,
      topRecommendation
        ? `\nSuggestion: ${topRecommendation.area}.${topRecommendation.parameter} ${topRecommendation.baseline || "n/a"} -> ${topRecommendation.suggested || "n/a"} (${topRecommendation.direction || "n/a"})`
        : "",
      pageUrl ? `\nStatus page: ${pageUrl}` : "",
      issueUrl ? `Live issue: ${issueUrl}` : "",
      runUrl ? `Run: ${runUrl}` : "",
    ].filter(Boolean).join("\n"),
    3500,
  );
}

async function sendWecomNotification(config, record, standingReport, options = {}) {
  if (!shouldSendWecom(config, options)) {
    return false;
  }

  const cfg = config.notifications?.wecom || {};
  const webhook = optionalEnv(cfg.webhook_url_env || "WECOM_WEBHOOK_URL", "");

  if (!webhook) {
    console.log("WeCom notification skipped because WECOM_WEBHOOK_URL is not set.");
    return false;
  }

  const content = buildWecomMarkdown(config, record, standingReport, options);
  const response = await fetch(webhook, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      msgtype: "markdown",
      markdown: { content },
    }),
  });
  const text = await response.text();

  if (!response.ok) {
    throw new Error(`WeCom webhook failed: HTTP ${response.status} ${text}`);
  }

  let result = {};
  try {
    result = JSON.parse(text);
  } catch {
    result = { raw: text };
  }

  if (result.errcode && result.errcode !== 0) {
    throw new Error(`WeCom webhook rejected message: ${text}`);
  }

  console.log("WeCom notification sent.");
  return true;
}

function appendStepSummary(markdown) {
  if (!process.env.GITHUB_STEP_SUMMARY || !markdown) {
    return false;
  }

  fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${markdown}\n`, "utf8");
  return true;
}

async function githubApi(config, method, endpoint, body) {
  const token = optionalEnv("GITHUB_TOKEN", "");
  const repository = optionalEnv("GITHUB_REPOSITORY", "");

  if (!token || !repository) {
    return null;
  }

  const response = await fetch(`https://api.github.com/repos/${repository}${endpoint}`, {
    method,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "User-Agent": "mgt267-watchdog",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`GitHub API ${method} ${endpoint} failed: ${response.status} ${text}`);
  }

  if (response.status === 204) {
    return {};
  }

  return response.json();
}

async function publishGithubIssueStatus(config, markdown) {
  const cfg = config.notifications?.github_issue || {};

  if (cfg.enabled === false || !isTruthyEnv("GITHUB_ISSUE_NOTIFICATIONS")) {
    return false;
  }

  const token = optionalEnv("GITHUB_TOKEN", "");
  const repository = optionalEnv("GITHUB_REPOSITORY", "");

  if (!token || !repository) {
    console.log("GitHub issue status skipped because GITHUB_TOKEN/GITHUB_REPOSITORY is unavailable.");
    return false;
  }

  const title = cfg.title || "MGT267 Watchdog Live Status";
  const body = [
    markdown,
    "---",
    "This issue is automatically updated by GitHub Actions. SMTP email can remain disabled while monitoring continues.",
  ].join("\n");
  const issues = await githubApi(config, "GET", "/issues?state=open&per_page=100");
  const exactMatches = (issues || []).filter(
    (issue) => !issue.pull_request && issue.title === title,
  );
  exactMatches.sort((left, right) => {
    const rightTime = new Date(right.updated_at || right.created_at || 0).getTime();
    const leftTime = new Date(left.updated_at || left.created_at || 0).getTime();
    return rightTime - leftTime || right.number - left.number;
  });
  const existing = exactMatches[0];

  if (existing) {
    await githubApi(config, "PATCH", `/issues/${existing.number}`, { body });
    for (const duplicate of exactMatches.slice(1)) {
      await githubApi(config, "PATCH", `/issues/${duplicate.number}`, {
        state: "closed",
        state_reason: "not_planned",
        body: `${body}\n\nDuplicate live status issue closed automatically; current issue is #${existing.number}.`,
      });
      console.log(`Duplicate GitHub live status issue closed: #${duplicate.number}`);
    }
    console.log(`GitHub live status issue updated: #${existing.number}`);
    return existing.html_url || "";
  }

  const created = await githubApi(config, "POST", "/issues", { title, body });
  console.log(`GitHub live status issue created: #${created.number}`);
  return created.html_url || "";
}

async function publishMonitorStatus(config, record, standingReport, options = {}) {
  let issueUrl = "";
  let githubIssueUpdated = false;

  try {
    issueUrl = (await publishGithubIssueStatus(
      config,
      buildWatchdogMarkdownSummary(config, record, standingReport, options),
    )) || "";
    githubIssueUpdated = Boolean(issueUrl);
  } catch (error) {
    console.log(`GitHub issue status failed: ${error.message}`);
  }

  const finalOptions = { ...options, githubIssueUrl: issueUrl };
  const markdown = buildWatchdogMarkdownSummary(
    config,
    record,
    standingReport,
    finalOptions,
  );
  const siteDir = writeStatusSite(
    config,
    record,
    standingReport,
    finalOptions,
    markdown,
  );
  console.log(`Status site: ${siteDir}`);
  appendStepSummary(markdown);

  try {
    await sendWecomNotification(config, record, standingReport, finalOptions);
  } catch (error) {
    console.log(`WeCom notification failed: ${error.message}`);
  }

  return githubIssueUpdated;
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

function buildAdjustmentPlanLines(plan) {
  const recommendations = plan?.recommendations || [];

  if (recommendations.length === 0) {
    return [];
  }

  return [
    "",
    "Auto-Adjustment Research Plan",
    `Mode: ${plan.mode || "research_only"} | submit_allowed: no`,
    `Posture: ${plan.posture || "n/a"}`,
    ...(plan.safety?.note ? [`Safety: ${plan.safety.note}`] : []),
    ...recommendations.map((item) =>
      [
        item.area,
        item.parameter,
        `baseline ${item.baseline || "n/a"}`,
        `suggested ${item.suggested || "n/a"}`,
        `change ${item.change || "n/a"}`,
        item.direction,
        item.urgency,
        item.reason,
      ].join(" | "),
    ),
  ];
}

function buildPolicyApplyLines(config, plan, policySnapshot = []) {
  if (config.policy_apply?.enabled === false) {
    return [];
  }

  const workflowUrl = policyApplyWorkflowUrl(config);
  const { changes, conflicts } = candidatePolicyChangesFromPlan(
    config,
    plan,
    policySnapshot,
    { allowShipping: false },
  );

  return [
    "",
    "Approval-Gated Apply",
    workflowUrl ? `Workflow: ${workflowUrl}` : "Workflow: not configured",
    "To apply: open workflow, choose recommended mode, type APPLY in confirm. The workflow re-crawls latest values and uses the latest safe recommended values before submitting.",
    "Optional edits: choose custom mode and fill only fields you want to override; blank custom fields still use the latest recommendation.",
    "Dry run: leave confirm as DRY_RUN or set dry_run=true.",
    "Guardrails: max +/-25 per numeric field, no priority/capacity changes, no decrease during lost demand or below dynamic cover minimum.",
    conflicts.length ? `Conflicts: ${conflicts.length} candidate conflict(s), apply workflow will block conflicting fields.` : "Conflicts: none",
    changes.length ? "Recommended apply fields:" : "Recommended apply fields: none",
    ...changes.map((item) =>
      [
        item.area,
        item.parameter,
        `override box ${policyWorkflowInputName(item.area, item.parameter) || "n/a"}`,
        `current ${item.current ?? "n/a"}`,
        `suggested ${item.suggested ?? "n/a"}`,
        item.direction,
        item.reason,
      ].join(" | "),
    ),
  ];
}

function buildPolicyLines(policySnapshot = []) {
  if (!policySnapshot.length) {
    return [];
  }

  const lines = ["", "Policy Snapshot"];

  for (const page of policySnapshot) {
    lines.push(`[${page.region || page.section}] ${page.label}`);

    for (const fact of (page.facts || []).slice(0, 3)) {
      lines.push(`- ${fact}`);
    }

    const controls = (page.forms || [])
      .filter((control) =>
        ["ship1", "point1", "quant1", "priority1", "policy"].includes(control.name),
      )
      .map((control) =>
        `${control.name}=${control.text || control.value || control.checked || "n/a"}`,
      );

    if (controls.length) {
      lines.push(`- ${controls.join(" | ")}`);
    }
  }

  return lines;
}

function buildBacktestLines(report) {
  if (!report?.enabled) {
    return [];
  }

  const latest = report.latest || {};
  const lines = [
    "",
    "Backtest",
    `Source: ${report.data_source}; local dependency: ${report.local_dependency}`,
    `Window: day ${report.day_start} to ${report.day_end}; observations=${report.observations}`,
    `Latest: inventory=${latest.warehouse_inventory ?? "n/a"} | demand=${latest.demand ?? "n/a"} | days_cover=${ratioMetric(latest.days_of_cover) || "n/a"} | shipment_demand=${ratioMetric(latest.shipment_to_demand_ratio) || "n/a"}`,
  ];

  for (const recommendation of report.recommendations || []) {
    lines.push(
      [
        recommendation.indicator,
        `${recommendation.operator} ${recommendation.suggested_threshold}`,
        `precision=${
          Number.isFinite(recommendation.precision)
            ? percentMetric(recommendation.precision * 100)
            : "n/a"
        }`,
        `recall=${
          Number.isFinite(recommendation.recall)
            ? percentMetric(recommendation.recall * 100)
            : "n/a"
        }`,
        `f1=${ratioMetric(recommendation.f1) || "n/a"}`,
      ].join(" | "),
    );
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

function formatMetricNumber(value, unit = "") {
  if (!Number.isFinite(value)) {
    return "";
  }

  if (unit === "%") {
    return `${value.toFixed(2)}%`;
  }

  if (unit === "ratio" || unit === "days") {
    return value.toFixed(2);
  }

  if (Math.abs(value) >= 1000) {
    return formatNumber(value, 2);
  }

  if (Number.isInteger(value)) {
    return String(value);
  }

  return value.toFixed(3).replace(/\.?0+$/, "");
}

function addMetric(metricMap, metric) {
  if (!metric || !metric.key) {
    return;
  }

  metricMap.set(metric.key, {
    unit: "",
    source: "crawler",
    ...metric,
    valueRaw:
      metric.valueRaw ?? formatMetricNumber(metric.valueNumber, metric.unit || ""),
  });
}

function nearestCompetitor(standingReport) {
  return standingReport.rows
    .filter((row) => row.team !== standingReport.target.team)
    .sort((a, b) => b.cashNumber - a.cashNumber)[0];
}

function buildMetricCatalog(config, record, standingReport, snapshot, options = {}) {
  const metricMap = new Map();

  addMetric(metricMap, {
    key: "dashboard:cash",
    label: "Dashboard cash",
    valueNumber: record.cashNumber,
    valueRaw: record.cash,
    unit: "$",
    source: "dashboard",
  });
  addMetric(metricMap, {
    key: "dashboard:day",
    label: "Dashboard day",
    valueNumber: Number(String(record.dashboardDay).replace(/,/g, "")),
    valueRaw: record.dashboardDay,
    unit: "day",
    source: "dashboard",
  });
  addMetric(metricMap, {
    key: "warehouse_inventory:warehouse",
    label: "Warehouse inventory",
    valueNumber: record.warehouseInventory,
    valueRaw: String(record.warehouseInventory),
    unit: "units",
    source: "warehouse inventory",
  });

  if (snapshot) {
    for (const section of snapshot.sections) {
      for (const plot of section.plots) {
        for (const series of plot.series) {
          addMetric(metricMap, {
            key: series.key,
            label: `${section.name} ${plot.label} ${series.series}`,
            valueNumber: series.valueNumber,
            valueRaw: series.valueRaw,
            unit: "units",
            source: `${section.name} ${plot.label}`,
            delta: series.delta,
            deltaText: series.deltaText,
            changePercent: series.changePercent,
            changePercentText: series.changePercentText,
          });
        }
      }
    }
  }

  const servedRegions = servedRegionsForWarehouse(
    options.policySnapshot || [],
    "warehouse_calopeia",
  );
  const demandRegions = servedRegions.length ? servedRegions : ["Calopeia"];
  const demand = sumMetricForRegions(metricMap, "hq_demand", demandRegions);
  const lostDemand = sumMetricForRegions(metricMap, "hq_lost_demand", demandRegions);
  const shipments = sumMetricForRegions(
    metricMap,
    "warehouse_shipments",
    demandRegions,
  );
  const wip = metricMap.get("factory_wip:Calopeia")?.valueNumber;
  const truckPipeline = metricMap.get("warehouse_inventory:truck")?.valueNumber;
  const mailPipeline = metricMap.get("warehouse_inventory:mail")?.valueNumber;
  const competitor = nearestCompetitor(standingReport);
  const coverTargets = coverageTargets(config, record);
  const rules = gameRules(config);

  addMetric(metricMap, {
    key: "derived:calopeia_served_region_count",
    label: "Calopeia served region count",
    valueNumber: demandRegions.length,
    valueRaw: String(demandRegions.length),
    unit: "regions",
    source: "warehouse policy",
  });
  addMetric(metricMap, {
    key: "derived:remaining_game_days",
    label: "Remaining game days",
    valueNumber: coverTargets.remainingDays,
    unit: "days",
    source: `game end day ${rules.end_day}`,
  });
  addMetric(metricMap, {
    key: "derived:cover_target_min",
    label: "Dynamic cover minimum",
    valueNumber: coverTargets.min,
    unit: "days",
    source: "configured target adjusted by remaining game days",
  });
  addMetric(metricMap, {
    key: "derived:cover_target_max",
    label: "Dynamic cover maximum",
    valueNumber: coverTargets.max,
    unit: "days",
    source: "configured target adjusted by remaining game days",
  });
  addMetric(metricMap, {
    key: "derived:cover_target",
    label: "Dynamic cover target",
    valueNumber: coverTargets.target,
    unit: "days",
    source: "configured target adjusted by remaining game days",
  });

  if (Number.isFinite(demand)) {
    addMetric(metricMap, {
      key: "derived:calopeia_served_demand",
      label: "Calopeia served demand",
      valueNumber: demand,
      unit: "units",
      source: `served regions: ${demandRegions.join(", ")}`,
    });
  }

  if (Number.isFinite(lostDemand)) {
    addMetric(metricMap, {
      key: "derived:calopeia_served_lost_demand",
      label: "Calopeia served lost demand",
      valueNumber: lostDemand,
      unit: "units",
      source: `served regions: ${demandRegions.join(", ")}`,
    });
  }

  if (Number.isFinite(shipments)) {
    addMetric(metricMap, {
      key: "derived:calopeia_served_shipments",
      label: "Calopeia served shipments",
      valueNumber: shipments,
      unit: "units",
      source: `served regions: ${demandRegions.join(", ")}`,
    });
  }

  if (Number.isFinite(demand) && demand > 0) {
    const daysOfCover = record.warehouseInventory / demand;
    const pipelineUnits = [wip, truckPipeline, mailPipeline]
      .filter(Number.isFinite)
      .reduce((sum, value) => sum + value, 0);
    const totalCover = (record.warehouseInventory + pipelineUnits) / demand;

    addMetric(metricMap, {
      key: "derived:days_of_cover",
      label: "Days of cover",
      valueNumber: daysOfCover,
      unit: "days",
      source: `derived from served demand: ${demandRegions.join(", ")}`,
    });
    addMetric(metricMap, {
      key: "derived:inbound_pipeline_units",
      label: "Inbound pipeline units",
      valueNumber: pipelineUnits,
      unit: "units",
      source: "factory WIP + warehouse mail + warehouse truck",
    });
    addMetric(metricMap, {
      key: "derived:pipeline_days_of_cover",
      label: "Inventory plus pipeline days of cover",
      valueNumber: totalCover,
      unit: "days",
      source: "warehouse inventory + factory/WIP transport pipeline",
    });
    addMetric(metricMap, {
      key: "derived:cover_shortage_gap",
      label: "Dynamic cover shortage gap",
      valueNumber: coverTargets.min - daysOfCover,
      unit: "days",
      source: "positive means served cover is below dynamic minimum",
    });
    addMetric(metricMap, {
      key: "derived:cover_excess_gap",
      label: "Dynamic cover excess gap",
      valueNumber: daysOfCover - coverTargets.max,
      unit: "days",
      source: "positive means served cover is above dynamic maximum",
    });
    addMetric(metricMap, {
      key: "derived:endgame_excess_cover_gap",
      label: "Endgame excess cover gap",
      valueNumber: Number.isFinite(coverTargets.remainingDays)
        ? daysOfCover - coverTargets.remainingDays
        : null,
      unit: "days",
      source: "positive means cover exceeds remaining demand horizon",
    });
  }

  if (Number.isFinite(lostDemand) && Number.isFinite(demand) && demand > 0) {
    addMetric(metricMap, {
      key: "derived:lost_demand_rate",
      label: "Lost demand rate",
      valueNumber: (lostDemand / demand) * 100,
      unit: "%",
      source: "derived",
    });
  }

  if (Number.isFinite(shipments) && Number.isFinite(demand) && demand > 0) {
    const shipmentRatio = shipments / demand;
    const shipmentRatioMin = Number(
      config.auto_adjust?.targets?.shipment_to_demand_ratio_min ?? 0.9,
    );

    addMetric(metricMap, {
      key: "derived:shipment_to_demand_ratio",
      label: "Shipment / demand ratio",
      valueNumber: shipmentRatio,
      unit: "ratio",
      source: "derived",
    });
    addMetric(metricMap, {
      key: "derived:shipment_shortage_gap",
      label: "Shipment coverage shortage gap",
      valueNumber: shipmentRatioMin - shipmentRatio,
      unit: "ratio",
      source: "positive means shipments are below target demand coverage",
    });
  }

  if (Number.isFinite(wip) && Number.isFinite(demand) && demand > 0) {
    addMetric(metricMap, {
      key: "derived:wip_to_demand_ratio",
      label: "WIP / demand ratio",
      valueNumber: wip / demand,
      unit: "ratio",
      source: "derived",
    });
  }

  if (competitor && Number.isFinite(competitor.cashNumber) && competitor.cashNumber !== 0) {
    addMetric(metricMap, {
      key: "derived:cash_lead_percent_vs_nearest",
      label: "Cash lead vs nearest competitor",
      valueNumber:
        ((standingReport.target.cashNumber - competitor.cashNumber) /
          competitor.cashNumber) *
        100,
      unit: "%",
      source: `standing vs ${competitor.team}`,
    });
  }

  return metricMap;
}

function compareRule(value, operator, threshold) {
  if (!Number.isFinite(value) || !Number.isFinite(threshold)) {
    return false;
  }

  switch (operator) {
    case ">":
    case "gt":
      return value > threshold;
    case ">=":
    case "gte":
      return value >= threshold;
    case "<":
    case "lt":
      return value < threshold;
    case "<=":
    case "lte":
      return value <= threshold;
    case "==":
    case "=":
    case "eq":
      return value === threshold;
    default:
      throw new Error(`Unsupported alert operator: ${operator}`);
  }
}

function legacyThresholdRules(config) {
  const thresholds = config.monitor.metric_thresholds || {};
  const rules = [];

  for (const [metric, threshold] of Object.entries(thresholds)) {
    if (!threshold || typeof threshold !== "object") {
      continue;
    }

    if (
      threshold.max !== null &&
      threshold.max !== undefined &&
      Number.isFinite(Number(threshold.max))
    ) {
      rules.push({
        id: `${metric}:max`,
        label: `${threshold.label || metric} max`,
        metric,
        operator: ">=",
        threshold: Number(threshold.max),
        severity: threshold.severity || "warning",
        channels: ["hourly"],
        message: "Legacy max threshold reached.",
      });
    }

    if (
      threshold.min !== null &&
      threshold.min !== undefined &&
      Number.isFinite(Number(threshold.min))
    ) {
      rules.push({
        id: `${metric}:min`,
        label: `${threshold.label || metric} min`,
        metric,
        operator: "<=",
        threshold: Number(threshold.min),
        severity: threshold.severity || "warning",
        channels: ["hourly"],
        message: "Legacy min threshold reached.",
      });
    }
  }

  return rules;
}

function buildWatchlist(config, metricMap, channel = "hourly") {
  const rules = [
    ...(config.monitor.alert_rules || []),
    ...legacyThresholdRules(config),
  ];

  return rules
    .filter((rule) => {
      if (rule.enabled === false) {
        return false;
      }

      const channels = rule.channels || ["hourly"];
      return channels.includes(channel);
    })
    .map((rule) => {
      const metric = metricMap.get(rule.metric);
      const threshold = Number(rule.threshold);
      const operator = rule.operator || rule.rule || ">=";
      const isAlert = metric
        ? compareRule(metric.valueNumber, operator, threshold)
        : false;

      return {
        id: rule.id,
        label: rule.label || rule.id,
        metricKey: rule.metric,
        metricLabel: metric?.label || rule.metric,
        current: metric?.valueNumber ?? null,
        currentRaw: metric?.valueRaw || "",
        unit: metric?.unit || "",
        operator,
        threshold,
        thresholdRaw: formatMetricNumber(threshold, metric?.unit || ""),
        severity: rule.severity || "warning",
        channels: rule.channels || ["hourly"],
        message: rule.message || "",
        status: metric ? (isAlert ? "ALERT" : "OK") : "NO DATA",
        isAlert,
      };
    });
}

function buildMetricAlertLines(watchlist = []) {
  const alerts = watchlist.filter((item) => item.isAlert);

  if (alerts.length === 0) {
    return [];
  }

  return [
    "",
    "[Alert Rules]",
    ...alerts.map((alert) =>
      [
        alert.label,
        `current ${alert.currentRaw}`,
        `${alert.operator} ${alert.thresholdRaw}`,
        alert.severity,
      ].join(" | "),
    ),
  ];
}

function buildWatchlistLines(watchlist = []) {
  if (watchlist.length === 0) {
    return [];
  }

  return [
    "",
    "[Watchlist]",
    ...watchlist.map((item) =>
      [
        item.status,
        item.severity,
        item.label,
        `current ${item.currentRaw || "n/a"}`,
        `${item.operator} ${item.thresholdRaw}`,
      ].join(" | "),
    ),
  ];
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
  const metricCatalog = options.metricCatalog;
  const wip = findOperationalMetric(snapshot, "factory_wip:Calopeia");
  const shipments = findOperationalMetric(snapshot, "warehouse_shipments:Calopeia");
  const servedRegions = servedRegionsForWarehouse(
    options.policySnapshot || [],
    "warehouse_calopeia",
  );
  const serviceRegions = servedRegions.length ? servedRegions : ["Calopeia"];
  const lostDemand =
    metricValue(metricCatalog, "derived:calopeia_served_lost_demand") ??
    metricValue(metricCatalog, "hq_lost_demand:Calopeia");
  const lostDemandRaw =
    metricRaw(metricCatalog, "derived:calopeia_served_lost_demand") ||
    metricRaw(metricCatalog, "hq_lost_demand:Calopeia") ||
    "n/a";
  const daysOfCover = metricValue(metricCatalog, "derived:days_of_cover");
  const shipmentRatio = metricValue(metricCatalog, "derived:shipment_to_demand_ratio");
  const targets = config.auto_adjust?.targets || {};
  const coverTargets = coverageTargets(config, record);
  const daysMin = coverTargets.min;
  const daysMax = coverTargets.max;
  const inventoryLow = Number(targets.warehouse_inventory_low ?? 50);
  const shipmentRatioMin = Number(targets.shipment_to_demand_ratio_min ?? 0.9);
  const rules = gameRules(config);
  const remainingDays = remainingGameDays(config, record);
  const policy = policyBaseline(config, options.policySnapshot || []);
  const factory = policy.factory || {};
  const factoryMethod = policyMethodText(factory);
  const factoryMethodLower = factoryMethod.toLowerCase();
  const factoryQuantity = Number(factory.quantity);
  const truckBreakEven =
    Number(rules.mail_cost_per_drum) > 0
      ? Math.ceil(Number(rules.truck_cost) / Number(rules.mail_cost_per_drum))
      : null;
  const daysCoverText = Number.isFinite(daysOfCover)
    ? `${formatMetricNumber(daysOfCover, "days")} days of cover`
    : "unknown days of cover";
  const shortageRisk =
    (Number.isFinite(lostDemand) && lostDemand > 0) ||
    (Number.isFinite(daysOfCover) && daysOfCover < daysMin) ||
    record.warehouseInventory <= inventoryLow;
  const highStock =
    (Number.isFinite(daysOfCover) && daysOfCover > daysMax);
  const checkpointOnly =
    record.inventoryCheckpoint &&
    serviceRegions.length > 1 &&
    Number.isFinite(daysOfCover) &&
    daysOfCover <= daysMax &&
    !shortageRisk;

  if (shortageRisk) {
    suggestions.push(
      `Shortage risk is active: inventory ${record.warehouseInventory}, ${daysCoverText}, served lost demand ${lostDemandRaw}; protect 24-hour customer fulfillment before cutting stock further.`,
    );
  } else if (highStock) {
    suggestions.push(
      `Inventory is high: warehouse ${record.warehouseInventory}, ${daysCoverText}; trim order point or batch quantity gradually because holding cost applies and unsold stock is worthless on day ${rules.end_day}.`,
    );
  } else if (checkpointOnly) {
    suggestions.push(
      `Do not reduce Calopeia just because it crossed the legacy 450 reference: it serves ${serviceRegions.join(", ")} and has ${daysCoverText}; hold or restore the current order point unless served lost demand appears.`,
    );
  } else {
    suggestions.push(
      `Inventory is within the decision band: warehouse ${record.warehouseInventory}, ${daysCoverText}; keep changes small and watch seasonal demand, shipments, and lost demand together.`,
    );
  }

  if (Number.isFinite(lostDemand) && lostDemand > 0) {
    suggestions.push(
      `Served lost demand is ${lostDemandRaw}; inspect warehouse availability and factory-to-warehouse lead time first, since unfilled orders are lost after 24 hours.`,
    );
  } else if (Number.isFinite(shipmentRatio) && shipmentRatio < shipmentRatioMin) {
    suggestions.push(
      `Shipments trail demand at ${formatMetricNumber(shipmentRatio, "ratio")}; verify whether warehouse stock or replenishment timing is constraining customer fulfillment.`,
    );
  }

  if (shortageRisk && factoryMethodLower.includes("truck")) {
    suggestions.push(
      `Factory-to-warehouse transport is currently ${factoryMethod}; mail costs more per large batch but arrives in ${rules.mail_lead_days} day versus truck in ${rules.truck_lead_days} days, so compare it for urgent recovery.`,
    );
  } else if (
    !shortageRisk &&
    Number.isFinite(factoryQuantity) &&
    Number.isFinite(truckBreakEven) &&
    factoryQuantity >= truckBreakEven &&
    factoryMethodLower.includes("mail")
  ) {
    suggestions.push(
      `Factory batch quantity ${factoryQuantity} is above the truck cost break-even near ${truckBreakEven} drums; if coverage is healthy, compare truck savings against the ${rules.truck_lead_days}-day delay.`,
    );
  } else if (wip?.valueNumber > 0 && shipments?.valueNumber === 0) {
    suggestions.push(
      `Factory WIP is ${wip.valueRaw} and shipments are ${shipments.valueRaw}; review batch completion and transport timing before starting more production.`,
    );
  }

  if (
    Number.isFinite(remainingDays) &&
    remainingDays <= Number(rules.capacity_expansion_lead_days) + 30
  ) {
    suggestions.push(
      `Only ${remainingDays} days remain and capacity takes ${rules.capacity_expansion_lead_days} days to arrive; avoid new capacity unless the backtest shows persistent lost demand.`,
    );
  }

  if (standingReport.target.rank === 1) {
    suggestions.push(
      `Cash rank is 1 at ${record.targetCash}; protect the lead by avoiding excess inventory, underfilled truck cost, and avoidable missed demand.`,
    );
  } else {
    suggestions.push(
      `Cash rank is ${standingReport.target.rank}; close the gap by eliminating lost demand first, then reduce carrying and transport waste before considering capacity.`,
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
    .map((line) => line.replace(/^\s*(?:[-*]|\d+[.)])\s*/, "").trim())
    .filter(Boolean)
    .slice(0, limit);
}

function recommendationGuardrailContext(config, record, options = {}) {
  const metricCatalog = options.metricCatalog;
  const regions = servedRegionsForWarehouse(
    options.policySnapshot || [],
    "warehouse_calopeia",
  );
  const serviceRegions = regions.length ? regions : ["Calopeia"];
  const targets = config.auto_adjust?.targets || {};
  const coverTargets = coverageTargets(config, record);
  const daysMin = coverTargets.min;
  const daysMax = coverTargets.max;
  const inventoryLow = Number(targets.warehouse_inventory_low ?? 50);
  const daysOfCover = metricValue(metricCatalog, "derived:days_of_cover");
  const lostDemand =
    metricValue(metricCatalog, "derived:calopeia_served_lost_demand") ??
    metricValue(metricCatalog, "hq_lost_demand:Calopeia");
  const shortageRisk =
    (Number.isFinite(lostDemand) && lostDemand > 0) ||
    (Number.isFinite(daysOfCover) && daysOfCover < daysMin) ||
    record.warehouseInventory <= inventoryLow;
  const trueHighCover = Number.isFinite(daysOfCover) && daysOfCover > daysMax;
  const checkpointOnly =
    record.inventoryCheckpoint &&
    serviceRegions.length > 1 &&
    Number.isFinite(daysOfCover) &&
    daysOfCover <= daysMax &&
    !shortageRisk;

  return {
    serviceRegions,
    shortageRisk,
    trueHighCover,
    checkpointOnly,
  };
}

function violatesRecommendationGuardrails(item, context) {
  const text = String(item || "").toLowerCase();

  if (/\bpriority\b/.test(text)) {
    return true;
  }

  if (/nearest/.test(text) && /fulfill|policy|switch|change/.test(text)) {
    return true;
  }

  if (/cancel|uncheck|stop serving|serve fewer|disable/.test(text)) {
    return true;
  }

  if (/clear (?:the )?(?:critical )?alert|critical alert/.test(text)) {
    return true;
  }

  const reductionVerb = /reduce|decrease|lower|cut|trim|draw down|scale back/.test(text);
  const inventoryTarget =
    /inventory|stock|warehouse|order point|quantity|production|batch/.test(text);

  if (context.checkpointOnly && reductionVerb && inventoryTarget) {
    return true;
  }

  if (
    context.checkpointOnly &&
    /excess stock|excess inventory|future accumulation|prevent.*accumulation/.test(text)
  ) {
    return true;
  }

  if (!context.shortageRisk && /switch|change|use/.test(text) && /\bmail\b/.test(text)) {
    return true;
  }

  return false;
}

function applyRecommendationGuardrails(config, record, items, fallback, options = {}) {
  const context = recommendationGuardrailContext(config, record, options);
  const limit = Number(config.ai?.max_suggestions || 4);
  const accepted = [];
  const rejected = [];
  const seen = new Set();

  function add(item) {
    const normalized = String(item || "").trim();
    const key = normalized.toLowerCase();

    if (!normalized || seen.has(key) || accepted.length >= limit) {
      return;
    }

    accepted.push(normalized);
    seen.add(key);
  }

  for (const item of items || []) {
    if (violatesRecommendationGuardrails(item, context)) {
      rejected.push(item);
    } else {
      add(item);
    }
  }

  if (rejected.length || accepted.length === 0) {
    for (const item of fallback || []) {
      add(item);
    }
  }

  return {
    items: accepted.slice(0, limit),
    rejected,
  };
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
  const watchlist = options.watchlist || [];
  const language = config.ai?.recommendation_language || "English";
  const lines = [
    "You are advising a team in a supply chain simulation.",
    "Use only the data below. Do not invent missing values.",
    `Write ${language}. Return 2 to 4 concise action recommendations.`,
    "Each recommendation should be one short sentence, useful for operations decisions, and avoid formulas.",
    "Focus on practical levers: order point, order quantity, shipping method, capacity timing, cash protection, and lost-demand prevention.",
    "Do not recommend changing priority level because it does not affect this assignment.",
    "Do not recommend reducing inventory, order point, quantity, or production solely because inventory is near or above 450; 450 is a legacy reference, not an active alert threshold.",
    "If Calopeia is serving multiple regions with zero lost demand and normal served days of cover, recommend hold/restore settings and watch the truck pipeline instead of cutting.",
    "Return JSON only in this shape: {\"recommendations\":[\"...\"]}",
    "",
    "Game rules:",
    ...gameRuleContextLines(config, record),
    "",
    "Service and anti-bullwhip context:",
    ...serviceContextLines(
      config,
      record,
      options.metricCatalog,
      options.policySnapshot || [],
    ),
    "",
    "Current policy context:",
    ...policyContextLines(config, options.policySnapshot || []),
    "",
    `Target team: ${record.targetTeam}`,
    `Checked at: ${record.checkedAtLocal} ${config.crawl.timezone}`,
    `Cash: ${record.targetCash}`,
    `Rank: ${record.targetRank}`,
    `Dashboard day: ${record.dashboardDay}`,
    `Warehouse inventory: ${record.warehouseInventory}`,
    `Warehouse inventory day: ${record.warehouseDay}`,
    `Legacy inventory reference: ${record.threshold}`,
    `Inventory reference crossed: ${record.inventoryCheckpoint ? "yes" : "no"}`,
    `Active inventory alert: ${record.inventoryAlert ? "yes" : "no"}`,
    "",
    "Standing rows:",
    ...standingSummaryLines(standingReport),
  ];

  if (options.operationalSnapshot) {
    lines.push("", "Operational metrics:", ...metricSummaryLines(options.operationalSnapshot));
  }

  if (watchlist.length > 0) {
    lines.push(
      "",
      "Watchlist rules:",
      ...watchlist.map(
        (item) =>
          `${item.status}: ${item.label}, current=${item.currentRaw || "n/a"}, rule=${item.operator} ${item.thresholdRaw}, severity=${item.severity}`,
      ),
    );
  }

  if (metricAlerts.length > 0) {
    lines.push(
      "",
      "Configured metric alerts:",
      ...metricAlerts.map(
        (alert) =>
          `${alert.label}: current=${alert.currentRaw}, rule=${alert.operator} ${alert.thresholdRaw}, severity=${alert.severity}`,
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
    const guarded = applyRecommendationGuardrails(
      config,
      record,
      items,
      fallback,
      options,
    );

    if (guarded.items.length === 0) {
      throw new Error("Gemini returned no recommendation items");
    }

    if (guarded.rejected.length > 0) {
      console.warn(
        `Gemini recommendations filtered by guardrails: ${guarded.rejected.length}`,
      );
    }

    return {
      source:
        guarded.rejected.length > 0
          ? `Gemini ${model} + local guardrails`
          : `Gemini ${model}`,
      items: guarded.items,
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
  const hasAlert =
    record.inventoryAlert ||
    (options.metricAlerts || []).some((item) => item.severity === "critical");
  const alertPrefix = hasAlert ? "ALERT " : "";

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
  const gameEntryUrl = config.email?.game_entry_url || config.crawl.entry_url;

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
    `Legacy inventory reference: ${record.threshold}`,
    `Inventory reference crossed: ${record.inventoryCheckpoint ? "yes" : "no"}`,
    `Active inventory alert: ${record.inventoryAlert ? "YES" : "no"}`,
    `Game entry URL: ${gameEntryUrl}`,
    ...buildWatchlistLines(options.watchlist),
    "",
    "Recommendations",
    ...((options.recommendations?.items || []).map((item) => `- ${item}`)),
    `Source: ${options.recommendations?.source || "n/a"}`,
    ...buildAdjustmentPlanLines(options.adjustmentPlan),
    ...buildPolicyApplyLines(
      config,
      options.adjustmentPlan,
      options.policySnapshot || [],
    ),
    ...buildDayChangeReviewLines(options.dayChangeReview),
    ...buildMetricAlertLines(options.metricAlerts),
    ...buildBacktestLines(options.backtestReport),
    ...buildOperationalLines(options.operationalSnapshot),
    ...buildPolicyLines(options.policySnapshot),
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

function buildGameLinkHtml(config) {
  const url = config.email?.game_entry_url || config.crawl.entry_url;

  if (!url) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #bfdbfe;border-radius:8px;background:#eff6ff;">',
    "<tr>",
    '<td style="padding:14px 16px;">',
    '<div style="font-size:13px;color:#1e3a8a;font-weight:700;margin-bottom:4px;">Supply Chain Game</div>',
    `<div style="font-size:12px;color:#475569;">Open the game entry page to log in and adjust current operations.</div>`,
    "</td>",
    '<td style="padding:14px 16px;text-align:right;white-space:nowrap;">',
    `<a href="${escapeHtml(url)}" style="display:inline-block;background:#1d4ed8;color:#ffffff;text-decoration:none;border-radius:6px;padding:10px 14px;font-size:13px;font-weight:700;">Open Game</a>`,
    "</td>",
    "</tr>",
    "</table>",
    "</div>",
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

function buildPolicySnapshotHtml(policySnapshot = []) {
  if (!policySnapshot.length) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Policy Snapshot</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
    '<thead><tr style="background:#e2e8f0;color:#334155;">',
    '<th style="padding:9px;text-align:left;">Area</th>',
    '<th style="padding:9px;text-align:left;">Status / Policy</th>',
    '<th style="padding:9px;text-align:left;">Current Parameters</th>',
    "</tr></thead>",
    "<tbody>",
    ...policySnapshot.map((page) => {
      const controls = (page.forms || [])
        .filter((control) =>
          ["ship1", "point1", "quant1", "priority1", "policy"].includes(control.name),
        )
        .map((control) =>
          `${control.name}: ${control.text || control.value || control.checked || "n/a"}`,
        );
      const factText = (page.facts || []).slice(0, 2).join(" ");

      return [
        "<tr>",
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:700;">${escapeHtml(`${page.region || page.section} ${page.label}`)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(factText || "n/a")}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(controls.join(" | ") || "n/a")}</td>`,
        "</tr>",
      ].join("");
    }),
    "</tbody>",
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

function buildAdjustmentPlanHtml(plan) {
  const recommendations = plan?.recommendations || [];

  if (recommendations.length === 0) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Auto-Adjustment Research Plan</div>',
    '<div style="border:1px solid #fed7aa;border-radius:8px;background:#fff7ed;padding:10px 12px;margin-bottom:10px;color:#9a3412;font-size:13px;line-height:1.45;">',
    `<strong>Research-only:</strong> ${escapeHtml(plan.safety?.note || "No game forms are submitted.")}`,
    "</div>",
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
    '<thead><tr style="background:#e2e8f0;color:#334155;">',
    '<th style="padding:9px;text-align:left;">Area</th>',
    '<th style="padding:9px;text-align:left;">Parameter</th>',
    '<th style="padding:9px;text-align:right;">Baseline</th>',
    '<th style="padding:9px;text-align:right;">Suggested</th>',
    '<th style="padding:9px;text-align:right;">Change</th>',
    '<th style="padding:9px;text-align:left;">Direction</th>',
    '<th style="padding:9px;text-align:left;">Reason</th>',
    "</tr></thead>",
    "<tbody>",
    ...recommendations.map((item) => {
      const directionColor =
        item.direction === "increase"
          ? "#047857"
          : item.direction === "decrease"
            ? "#b91c1c"
            : "#334155";

      return [
        "<tr>",
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:700;">${escapeHtml(item.area)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(item.parameter)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(item.baseline || "n/a")}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;">${escapeHtml(item.suggested || "n/a")}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(item.change || "n/a")}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;color:${directionColor};font-weight:700;">${escapeHtml(item.direction)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(item.reason)}</td>`,
        "</tr>",
      ].join("");
    }),
    "</tbody>",
    "</table>",
    "</div>",
  ].join("");
}

function buildPolicyApplyHtml(config, plan, policySnapshot = []) {
  if (config.policy_apply?.enabled === false) {
    return "";
  }

  const workflowUrl = policyApplyWorkflowUrl(config);
  const { changes, conflicts } = candidatePolicyChangesFromPlan(
    config,
    plan,
    policySnapshot,
    { allowShipping: false },
  );

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Approval-Gated Apply</div>',
    '<div style="border:1px solid #bbf7d0;border-radius:8px;background:#f0fdf4;padding:12px 14px;margin-bottom:10px;color:#14532d;font-size:13px;line-height:1.45;">',
    '<strong>Semi-automatic only:</strong> this email does not change the game. The GitHub workflow re-crawls latest data, applies guardrails, and submits only after manual confirmation.',
    "</div>",
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin-bottom:10px;">',
    "<tr>",
    '<td style="font-size:13px;color:#334155;line-height:1.45;">',
    '<div><strong>One-click default:</strong> mode=recommended, confirm=APPLY. The workflow fills the latest safe recommended values at runtime.</div>',
    '<div><strong>Optional edits:</strong> choose mode=custom and fill only fields you want to override; blank fields still use the latest recommendation.</div>',
    '<div><strong>Dry run:</strong> confirm=DRY_RUN or dry_run=true</div>',
    '<div><strong>Guardrails:</strong> max +/-25 per numeric field, no priority/capacity changes, no decrease during lost demand or below dynamic cover minimum.</div>',
    conflicts.length
      ? `<div style="color:#b91c1c;font-weight:700;">Conflicts: ${escapeHtml(conflicts.length)} candidate conflict(s); conflicting fields are blocked.</div>`
      : '<div style="color:#047857;font-weight:700;">Conflicts: none</div>',
    "</td>",
    '<td style="text-align:right;vertical-align:top;white-space:nowrap;">',
    workflowUrl
      ? `<a href="${escapeHtml(workflowUrl)}" style="display:inline-block;background:#047857;color:#ffffff;text-decoration:none;border-radius:6px;padding:10px 14px;font-size:13px;font-weight:700;">Open Apply Workflow</a>`
      : "",
    "</td>",
    "</tr>",
    "</table>",
    changes.length
      ? [
          '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
          '<thead><tr style="background:#dcfce7;color:#14532d;">',
          '<th style="padding:9px;text-align:left;">Area</th>',
          '<th style="padding:9px;text-align:left;">Field</th>',
          '<th style="padding:9px;text-align:left;">Override Box</th>',
          '<th style="padding:9px;text-align:right;">Current</th>',
          '<th style="padding:9px;text-align:right;">Suggested</th>',
          '<th style="padding:9px;text-align:left;">Reason</th>',
          "</tr></thead>",
          "<tbody>",
          ...changes.map((item) =>
            [
              "<tr>",
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:700;">${escapeHtml(item.area)}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(item.parameter)}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-family:Consolas,monospace;font-size:12px;">${escapeHtml(policyWorkflowInputName(item.area, item.parameter) || "n/a")}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(item.current ?? "n/a")}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;">${escapeHtml(item.suggested ?? "n/a")}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(item.reason)}</td>`,
              "</tr>",
            ].join(""),
          ),
          "</tbody>",
          "</table>",
        ].join("")
      : '<div style="border:1px solid #d8dee9;border-radius:8px;background:#f8fafc;padding:10px 12px;color:#64748b;font-size:13px;">No policy fields are currently eligible for semi-automatic apply.</div>',
    "</div>",
  ].join("");
}

function statusColor(status) {
  if (status === "ALERT") {
    return "#b91c1c";
  }

  if (status === "NO DATA") {
    return "#64748b";
  }

  return "#047857";
}

function buildWatchlistHtml(watchlist = []) {
  if (watchlist.length === 0) {
    return "";
  }

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Decision Watchlist</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
    '<thead><tr style="background:#e2e8f0;color:#334155;">',
    '<th style="padding:9px;text-align:left;">Status</th>',
    '<th style="padding:9px;text-align:left;">Severity</th>',
    '<th style="padding:9px;text-align:left;">Indicator</th>',
    '<th style="padding:9px;text-align:right;">Current</th>',
    '<th style="padding:9px;text-align:right;">Rule</th>',
    "</tr></thead>",
    "<tbody>",
    ...watchlist.map((item) => {
      const color = statusColor(item.status);
      const background = item.isAlert ? "#fff7ed" : "#ffffff";

      return [
        `<tr style="background:${background};">`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;color:${color};font-weight:700;">${escapeHtml(item.status)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(item.severity)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:600;">${escapeHtml(item.label)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(item.currentRaw || "n/a")}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(`${item.operator} ${item.thresholdRaw}`)}</td>`,
        "</tr>",
      ].join("");
    }),
    "</tbody>",
    "</table>",
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
        `<td style="padding:8px;border-bottom:1px solid #fecaca;text-align:right;color:#b91c1c;font-weight:700;">${escapeHtml(`${alert.operator} ${alert.thresholdRaw}`)}</td>`,
        "</tr>",
      ].join(""),
    ),
    "</tbody>",
    "</table>",
    "</div>",
  ].join("");
}

function buildBacktestHtml(report) {
  if (!report?.enabled) {
    return "";
  }

  const latest = report.latest || {};
  const recommendations = report.recommendations || [];

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Backtest Snapshot</div>',
    '<div style="border:1px solid #d8dee9;border-radius:8px;background:#f8fafc;padding:10px 12px;margin-bottom:10px;color:#334155;font-size:13px;line-height:1.45;">',
    `Source: ${escapeHtml(report.data_source)}; local dependency: ${escapeHtml(report.local_dependency)}. `,
    `Window: day ${escapeHtml(report.day_start)} to ${escapeHtml(report.day_end)} (${escapeHtml(report.observations)} observations). `,
    `Latest inventory ${escapeHtml(latest.warehouse_inventory ?? "n/a")}, demand ${escapeHtml(latest.demand ?? "n/a")}, days cover ${escapeHtml(ratioMetric(latest.days_of_cover) || "n/a")}.`,
    "</div>",
    recommendations.length
      ? [
          '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;">',
          '<thead><tr style="background:#e2e8f0;color:#334155;">',
          '<th style="padding:9px;text-align:left;">Indicator</th>',
          '<th style="padding:9px;text-align:right;">Threshold</th>',
          '<th style="padding:9px;text-align:right;">Precision</th>',
          '<th style="padding:9px;text-align:right;">Recall</th>',
          '<th style="padding:9px;text-align:right;">F1</th>',
          "</tr></thead>",
          "<tbody>",
          ...recommendations.map((item) =>
            [
              "<tr>",
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:700;">${escapeHtml(item.indicator)}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(`${item.operator} ${item.suggested_threshold}`)}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(Number.isFinite(item.precision) ? percentMetric(item.precision * 100) : "n/a")}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(Number.isFinite(item.recall) ? percentMetric(item.recall * 100) : "n/a")}</td>`,
              `<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">${escapeHtml(ratioMetric(item.f1) || "n/a")}</td>`,
              "</tr>",
            ].join(""),
          ),
          "</tbody>",
          "</table>",
        ].join("")
      : "",
    "</div>",
  ].join("");
}

function buildDayChangeReviewHtml(review) {
  if (!review?.enabled) {
    return "";
  }

  const changeText = review.changed
    ? `changed from ${review.previous_day ?? "n/a"} to ${review.current_day ?? "n/a"}`
    : `unchanged at ${review.current_day ?? "n/a"}`;
  const rows = [
    ["Day status", changeText],
    ["Game clock", `day ${review.current_day ?? "n/a"} / ${review.end_day}; remaining ${review.remaining_days ?? "n/a"}; phase ${review.phase || "n/a"}`],
    ["Posture", review.posture || "n/a"],
    ["Latest operations", `inventory ${review.latest?.warehouse_inventory ?? "n/a"}, demand ${review.latest?.demand ?? "n/a"}, shipments ${review.latest?.shipments ?? "n/a"}, lost demand ${review.latest?.lost_demand ?? "n/a"}, cover ${review.latest?.days_of_cover ?? "n/a"}`],
    ["Cash lead", review.latest?.cash_lead_percent_vs_nearest || "n/a"],
    ["Backtest window", `${review.backtest?.window || "n/a"}; observations ${review.backtest?.observations ?? "n/a"}; horizon ${review.backtest?.horizon_days ?? "n/a"} days`],
  ];
  const actionItems = review.action_candidates?.length
    ? review.action_candidates
    : ["No policy changes recommended; hold current policy unless alerts worsen."];
  const backtestItems = review.backtest?.recommendations?.length
    ? review.backtest.recommendations
    : [`No strong historical threshold in ${review.backtest?.window || "the current window"}.`];

  return [
    '<div style="padding:0 22px 18px;">',
    '<div style="font-size:16px;font-weight:700;margin:4px 0 10px;color:#0f172a;">Simulated Day Review</div>',
    '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d8dee9;border-radius:8px;overflow:hidden;font-size:13px;margin-bottom:10px;">',
    '<tbody>',
    ...rows.map(([label, value]) =>
      [
        '<tr>',
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;background:#f8fafc;font-weight:700;width:28%;">${escapeHtml(label)}</td>`,
        `<td style="padding:8px;border-bottom:1px solid #e2e8f0;">${escapeHtml(value)}</td>`,
        '</tr>',
      ].join(""),
    ),
    '</tbody></table>',
    '<div style="border:1px solid #bfdbfe;border-radius:8px;background:#eff6ff;padding:10px 12px;margin-bottom:10px;color:#1e3a8a;font-size:13px;line-height:1.45;">',
    '<div style="font-weight:700;margin-bottom:5px;">Action candidates</div>',
    '<ol style="margin:0 0 0 20px;padding:0;">',
    ...actionItems.map((item) => `<li>${escapeHtml(item)}</li>`),
    '</ol></div>',
    '<div style="border:1px solid #d8dee9;border-radius:8px;background:#f8fafc;padding:10px 12px;color:#334155;font-size:13px;line-height:1.45;">',
    '<div style="font-weight:700;margin-bottom:5px;">Backtest and guardrails</div>',
    '<ul style="margin:0 0 8px 18px;padding:0;">',
    ...backtestItems.map((item) => `<li>${escapeHtml(item)}</li>`),
    '</ul>',
    '<ul style="margin:0 0 0 18px;padding:0;">',
    ...(review.guardrails || []).map((item) => `<li>${escapeHtml(item)}</li>`),
    '</ul></div>',
    '</div>',
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
  const alertCount = (options.metricAlerts || []).length;
  const hasCriticalAlert = (options.metricAlerts || []).some(
    (item) => item.severity === "critical",
  );
  const alertText =
    alertCount > 0
      ? `${alertCount} watchlist alert${alertCount === 1 ? "" : "s"} detected.`
      : "No configured watchlist alerts are active.";
  const bannerColor = isWarning || hasCriticalAlert ? "#b91c1c" : "#1d4ed8";
  const bannerBg = isWarning || hasCriticalAlert ? "#fef2f2" : "#eff6ff";

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
    buildCard("Reference", record.threshold, "#64748b"),
    buildCard("Report type", isWarning ? `${minutes}-minute warning` : "hourly"),
    "</tr>",
    "</table>",
    buildGameLinkHtml(config),
    buildDayChangeReviewHtml(options.dayChangeReview),
    buildWatchlistHtml(options.watchlist),
    buildRecommendationsHtml(options.recommendations),
    buildAdjustmentPlanHtml(options.adjustmentPlan),
    buildPolicyApplyHtml(
      config,
      options.adjustmentPlan,
      options.policySnapshot || [],
    ),
    buildMetricAlertsHtml(options.metricAlerts),
    buildBacktestHtml(options.backtestReport),
    !isWarning ? buildOperationalSnapshotHtml(options.operationalSnapshot) : "",
    !isWarning ? buildPolicySnapshotHtml(options.policySnapshot) : "",
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
  const normalizedRecipients = normalizeRecipients(recipients);

  if (normalizedRecipients.length === 0) {
    throw new Error("Email report is enabled, but email.recipients is empty");
  }

  const channel = options.kind === "warning" ? "warning" : "hourly";
  const metricCatalog =
    options.metricCatalog ??
    buildMetricCatalog(config, record, standingReport, options.operationalSnapshot, {
      policySnapshot: options.policySnapshot || [],
    });
  const watchlist = options.watchlist ?? buildWatchlist(config, metricCatalog, channel);
  const metricAlerts =
    options.metricAlerts ?? watchlist.filter((item) => item.isAlert);
  const recommendations =
    options.recommendations ??
    (await buildRecommendations(config, record, standingReport, {
      ...options,
      metricCatalog,
      watchlist,
      metricAlerts,
    }));
  const adjustmentPlan =
    options.adjustmentPlan ??
    buildAutoAdjustmentPlan(
      config,
      metricCatalog,
      record,
      standingReport,
      options.policySnapshot,
    );
  const renderOptions = {
    ...options,
    metricCatalog,
    watchlist,
    metricAlerts,
    recommendations,
    adjustmentPlan,
    policySnapshot: options.policySnapshot || [],
  };
  const subject = buildEmailSubject(config, record, renderOptions);
  const text = buildReportText(config, record, standingReport, renderOptions);
  const html = buildReportHtml(config, record, standingReport, renderOptions);
  const attachments = [];

  if (config.email.attach_excel !== false && options.plotSnapshots?.length) {
    const workbookBuffer = await buildDataWorkbookBuffer(
      config,
      record,
      standingReport,
      options.plotSnapshots,
      {
        metricCatalog,
        watchlist,
        operationalSnapshot: options.operationalSnapshot,
        adjustmentPlan,
        policySnapshot: options.policySnapshot || [],
        backtestReport: options.backtestReport,
      },
    );

    if (workbookBuffer) {
      attachments.push({
        filename: dataWorkbookFilename(record, options),
        content: workbookBuffer,
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
    }
  }

  if (isEmailDryRun()) {
    console.log("EMAIL_DRY_RUN=1, email report not sent.");
    console.log(`Dry-run recipients: ${normalizedRecipients.join(", ")}`);
    console.log(`Dry-run subject: ${subject}`);
    console.log(
      `Dry-run Excel attachment: ${
        attachments.length > 0 ? attachments[0].filename : "none"
      }`,
    );
    console.log(`Recommendations source: ${recommendations.source}`);
    for (const item of metricAlerts) {
      console.log(
        `Alert: ${item.label} current=${item.currentRaw || "n/a"} rule=${item.operator} ${item.thresholdRaw}`,
      );
    }
    for (const item of recommendations.items) {
      console.log(`- ${item}`);
    }
    for (const item of adjustmentPlan.recommendations || []) {
      console.log(
        `Adjustment research: ${item.area}.${item.parameter} ${item.baseline || "n/a"} -> ${item.suggested || "n/a"} (${item.direction})`,
      );
    }
    return false;
  }

  const transporter = createTransport();
  await transporter.verify();
  const deliveryLog = {
    checked_at: record.checkedAt,
    checked_at_local: record.checkedAtLocal,
    kind: options.kind || "hourly",
    subject,
    recipients: normalizedRecipients,
    attachment_filenames: attachments.map((attachment) => attachment.filename),
    results: [],
  };
  const baseMail = {
    from: formatFromAddress(),
    subject,
    text,
    html,
    attachments,
    headers: {
      "X-MGT267-Watchdog": "true",
      "X-MGT267-Report-Kind": options.kind || "hourly",
      "X-MGT267-Target-Team": record.targetTeam,
    },
  };
  const failures = [];

  for (const recipient of normalizedRecipients) {
    const info = await transporter.sendMail({
      ...baseMail,
      to: recipient,
    });
    const accepted = (info.accepted || []).map(String);
    const rejected = (info.rejected || []).map(String);
    const pending = (info.pending || []).map(String);
    const recipientAccepted = accepted.some(
      (acceptedRecipient) =>
        acceptedRecipient.toLowerCase() === recipient.toLowerCase(),
    );

    deliveryLog.results.push({
      recipient,
      accepted,
      rejected,
      pending,
      message_id: info.messageId || "",
      response: info.response || "",
      envelope: info.envelope || {},
    });
    console.log(
      `SMTP delivery recipient=${recipient} accepted=${accepted.join("|") || "none"} rejected=${rejected.join("|") || "none"} messageId=${info.messageId || "n/a"}`,
    );

    if (!recipientAccepted || rejected.length > 0) {
      failures.push(recipient);
    }
  }

  deliveryLog.success = failures.length === 0;
  deliveryLog.delivery_log_path = writeEmailDeliveryLog(config, deliveryLog);
  console.log(`Email delivery log: ${deliveryLog.delivery_log_path}`);

  if (failures.length > 0) {
    throw new Error(`SMTP did not accept all recipients: ${failures.join(", ")}`);
  }

  return true;
}

async function loginToGame(config) {
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

  return {
    cookieJar,
    dashboard,
    dashboardHtml,
  };
}

async function crawl(config, options = {}) {
  const session = await loginToGame(config);
  const { cookieJar, dashboard } = session;
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

  const policySnapshot = [];

  if (options.includePolicyPages !== false) {
    for (const page of config.crawl.policy_pages || []) {
      const pageResponse = await request(page.url, { method: "GET" }, cookieJar);
      const pageHtml = await pageResponse.text();

      if (!pageResponse.ok) {
        throw new Error(`${page.label} policy page failed with HTTP ${pageResponse.status}`);
      }

      policySnapshot.push(parsePolicyPageSnapshot(pageHtml, page));
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
    policySnapshot,
    standingReport,
    cookieJar: options.returnCookieJar ? cookieJar : undefined,
  };
}

function createRecord(config, dashboard, inventoryTable, standingReport, options = {}) {
  const checkedAt = options.checkedAt || new Date().toISOString();
  const threshold = Number(config.monitor.warehouse_inventory_threshold);
  const thresholdEnabled = config.monitor.warehouse_inventory_threshold_enabled === true;
  const currentInventory = inventoryTable.latestWarehouse.inventory;
  const inventoryCheckpoint =
    Number.isFinite(threshold) && currentInventory >= threshold;
  const inventoryAlert =
    options.inventoryAlertOverride ?? (thresholdEnabled && inventoryCheckpoint);

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
    inventoryCheckpoint,
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
  const policySnapshotCsvPath = path.resolve(
    process.cwd(),
    config.output.policy_snapshot_csv ||
      ".monitor-state/policy_snapshot_latest.csv",
  );
  const adjustmentPlanJsonPath = path.resolve(
    process.cwd(),
    config.output.adjustment_plan_json ||
      ".monitor-state/adjustment_plan_latest.json",
  );
  const adjustmentPlanCsvPath = path.resolve(
    process.cwd(),
    config.output.adjustment_plan_csv ||
      ".monitor-state/adjustment_plan_latest.csv",
  );
  const dataWorkbookPath = path.resolve(
    process.cwd(),
    config.output.data_workbook_xlsx,
  );
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));

  const previousState = readJson(statePath, {});
  const { dashboard, inventoryTable, plotSnapshots, standingReport, policySnapshot } =
    await crawl(config);
  const operationalSnapshot = buildOperationalSnapshot(
    plotSnapshots,
    previousState.last_operational_metrics || {},
  );
  const checkedAt = new Date().toISOString();
  const record = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });
  const metricCatalog = buildMetricCatalog(
    config,
    record,
    standingReport,
    operationalSnapshot,
    { policySnapshot },
  );
  const watchlist = buildWatchlist(config, metricCatalog, "hourly");
  const adjustmentPlan = buildAutoAdjustmentPlan(
    config,
    metricCatalog,
    record,
    standingReport,
    policySnapshot,
  );
  const backtestReport = buildBacktestReport(
    config,
    record,
    standingReport,
    plotSnapshots,
    policySnapshot,
  );
  const dayChangeReview = buildDayChangeReview(
    config,
    previousState,
    record,
    metricCatalog,
    adjustmentPlan,
    backtestReport,
  );
  const backtestOutputs = await writeBacktestOutputs(config, backtestReport);
  let emailSent = false;
  let emailError = "";
  const reportDecision = shouldSendReportNow(config, previousState);

  if (config.monitor.send_report_every_run && reportDecision.send) {
    try {
      emailSent = await sendReportEmail(config, record, standingReport, {
        operationalSnapshot,
        plotSnapshots,
        metricCatalog,
        watchlist,
        adjustmentPlan,
        policySnapshot,
        backtestReport,
        dayChangeReview,
      });
    } catch (error) {
      emailError = error.message;
      console.log(`Email report failed but monitoring will continue: ${emailError}`);
    }
  } else if (config.monitor.send_report_every_run) {
    console.log(`Email report skipped: ${reportDecision.reason}`);
  }

  fs.writeFileSync(inventoryCsvPath, buildWarehouseCsv(inventoryTable), "utf8");
  fs.writeFileSync(standingCsvPath, buildStandingGapsCsv(standingReport), "utf8");
  fs.writeFileSync(
    operationalCsvPath,
    buildOperationalSnapshotCsv(operationalSnapshot),
    "utf8",
  );
  fs.writeFileSync(
    policySnapshotCsvPath,
    buildPolicySnapshotCsv(policySnapshot),
    "utf8",
  );
  fs.writeFileSync(
    adjustmentPlanJsonPath,
    `${JSON.stringify(adjustmentPlan, null, 2)}\n`,
    "utf8",
  );
  fs.writeFileSync(
    adjustmentPlanCsvPath,
    buildAdjustmentPlanCsv(adjustmentPlan),
    "utf8",
  );
  const workbookWritten = await writeDataWorkbookFile(
    config,
    record,
    standingReport,
    plotSnapshots,
    dataWorkbookPath,
    {
      operationalSnapshot,
      metricCatalog,
      watchlist,
      adjustmentPlan,
      policySnapshot,
      backtestReport,
    },
  );
  const metricAlerts = watchlist.filter((item) => item.isAlert);
  await publishMonitorStatus(config, record, standingReport, {
    kind: "hourly",
    operationalSnapshot,
    metricCatalog,
    watchlist,
    metricAlerts,
    adjustmentPlan,
    policySnapshot,
    backtestReport,
    dayChangeReview,
    emailSent,
    emailError,
  });
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
        last_dashboard_day_number: dashboardDayNumber(record),
        last_warehouse_inventory: record.warehouseInventory,
        last_warehouse_day: inventoryTable.latestWarehouse.day,
        last_target_team: standingReport.target.team,
        last_target_rank: standingReport.target.rank,
        last_target_cash: standingReport.target.cash,
        last_target_cash_number: standingReport.target.cashNumber,
        last_threshold: record.threshold,
        last_inventory_alert: record.inventoryAlert,
        last_alerts: watchlist.filter((item) => item.isAlert),
        last_operational_metrics: operationalSnapshot.metrics,
        last_adjustment_plan: adjustmentPlan,
        last_backtest: {
          generated_at: backtestReport.generated_at,
          data_source: backtestReport.data_source,
          local_dependency: backtestReport.local_dependency,
          observations: backtestReport.observations,
          day_start: backtestReport.day_start,
          day_end: backtestReport.day_end,
          latest: backtestReport.latest,
          recommendations: backtestReport.recommendations,
        },
        last_day_change_review_at: dayChangeReview.changed
          ? checkedAt
          : previousState.last_day_change_review_at || null,
        last_day_change_dashboard_day: dayChangeReview.changed
          ? record.dashboardDay
          : previousState.last_day_change_dashboard_day || null,
        last_day_change_review: dayChangeReview,
        last_policy_pages: policySnapshot,
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
  console.log(`Sim day changed: ${dayChangeReview.changed ? "yes" : "no"}`);
  console.log(`Previous dashboard day: ${dayChangeReview.previous_day ?? "n/a"}`);
  console.log(`Remaining game days: ${dayChangeReview.remaining_days ?? "n/a"}`);
  console.log(`Target team: ${standingReport.target.team}`);
  console.log(`Target rank: ${standingReport.target.rank}`);
  console.log(`Target cash: ${standingReport.target.cash}`);
  console.log(`Warehouse inventory: ${record.warehouseInventory}`);
  console.log(`Warehouse inventory day: ${inventoryTable.latestWarehouse.day}`);
  console.log(`Legacy inventory reference: ${record.threshold}`);
  console.log(`Inventory reference crossed: ${record.inventoryCheckpoint ? "yes" : "no"}`);
  console.log(`Active inventory alert: ${record.inventoryAlert ? "yes" : "no"}`);
  console.log(`Report decision: ${reportDecision.reason}`);
  console.log(`Email sent: ${emailSent ? "yes" : "no"}`);
  console.log(`State: ${statePath}`);
  console.log(`History: ${historyPath}`);
  console.log(`Standing gaps: ${standingCsvPath}`);
  console.log(`Operational snapshot: ${operationalCsvPath}`);
  console.log(`Policy snapshot: ${policySnapshotCsvPath}`);
  console.log(`Adjustment plan: ${adjustmentPlanJsonPath}`);
  console.log(`Backtest JSON: ${backtestOutputs.jsonPath}`);
  console.log(`Backtest summary: ${backtestOutputs.summaryCsvPath}`);
  console.log(`Backtest workbook: ${backtestOutputs.workbookPath || "not written"}`);
  console.log(
    `Excel data workbook: ${workbookWritten ? dataWorkbookPath : "not written"}`,
  );
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
  const { dashboard, inventoryTable, plotSnapshots, standingReport, policySnapshot } =
    await crawl(config);
  const operationalSnapshot = buildOperationalSnapshot(
    plotSnapshots,
    previousState.last_operational_metrics || {},
  );
  const checkedAt = new Date().toISOString();
  const baseRecord = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });
  const backtestReport = buildBacktestReport(
    config,
    baseRecord,
    standingReport,
    plotSnapshots,
    policySnapshot,
  );
  const backtestOutputs = await writeBacktestOutputs(config, backtestReport);

  const hourlySent = await sendReportEmail(config, baseRecord, standingReport, {
    kind: "hourly",
    test: true,
    operationalSnapshot,
    plotSnapshots,
    policySnapshot,
    backtestReport,
  });

  console.log("Test email summary:");
  console.log(`Hourly report email sent: ${hourlySent ? "yes" : "no"}`);
  console.log(`Target team: ${baseRecord.targetTeam}`);
  console.log(`Target cash: ${baseRecord.targetCash}`);
  console.log(`Dashboard day: ${baseRecord.dashboardDay}`);
  console.log(`Warehouse inventory: ${baseRecord.warehouseInventory}`);
  console.log(`Operational metrics: ${Object.keys(operationalSnapshot.metrics).length}`);
  console.log(`Backtest workbook: ${backtestOutputs.workbookPath || "not written"}`);
}

async function sendWarningEmail(config) {
  const warningMinutes = Number(config.monitor.warning_minutes || 15);
  const statePath = path.resolve(process.cwd(), config.output.latest_json);
  const dataWorkbookPath = path.resolve(
    process.cwd(),
    config.output.data_workbook_xlsx,
  );
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));
  const previousState = readJson(statePath, {});
  const { dashboard, inventoryTable, plotSnapshots, standingReport, policySnapshot } =
    await crawl(config);
  const operationalSnapshot = buildOperationalSnapshot(
    plotSnapshots,
    previousState.last_operational_metrics || {},
  );
  const record = createRecord(config, dashboard, inventoryTable, standingReport);
  const metricCatalog = buildMetricCatalog(
    config,
    record,
    standingReport,
    operationalSnapshot,
    { policySnapshot },
  );
  const watchlist = buildWatchlist(config, metricCatalog, "warning");
  const warningAlerts = watchlist.filter((item) => item.isAlert);
  const adjustmentPlan = buildAutoAdjustmentPlan(
    config,
    metricCatalog,
    record,
    standingReport,
    policySnapshot,
  );
  const backtestReport = buildBacktestReport(
    config,
    record,
    standingReport,
    plotSnapshots,
    policySnapshot,
  );
  const dayChangeReview = buildDayChangeReview(
    config,
    previousState,
    record,
    metricCatalog,
    adjustmentPlan,
    backtestReport,
  );
  const backtestOutputs = await writeBacktestOutputs(config, backtestReport);
  const workbookWritten = await writeDataWorkbookFile(
    config,
    record,
    standingReport,
    plotSnapshots,
    dataWorkbookPath,
    {
      operationalSnapshot,
      metricCatalog,
      watchlist,
      adjustmentPlan,
      policySnapshot,
      backtestReport,
    },
  );
  let emailSent = false;
  let emailError = "";

  if (warningAlerts.length > 0 || dayChangeReview.changed) {
    try {
      emailSent = await sendReportEmail(
        config,
        record,
        standingReport,
        {
          kind: "warning",
          warningMinutes,
          operationalSnapshot,
          plotSnapshots,
          policySnapshot,
          metricCatalog,
          watchlist,
          metricAlerts: warningAlerts,
          adjustmentPlan,
          backtestReport,
          dayChangeReview,
        },
      );
    } catch (error) {
      emailError = error.message;
      console.log(`Warning email failed but monitoring will continue: ${emailError}`);
    }
  }

  await publishMonitorStatus(config, record, standingReport, {
    kind: "warning",
    warningMinutes,
    operationalSnapshot,
    plotSnapshots,
    policySnapshot,
    metricCatalog,
    watchlist,
    metricAlerts: warningAlerts,
    adjustmentPlan,
    backtestReport,
    dayChangeReview,
    emailSent,
    emailError,
  });
  fs.writeFileSync(
    statePath,
    `${JSON.stringify(
      {
        ...previousState,
        last_run_at: record.checkedAt,
        last_run_at_local: record.checkedAtLocal,
        last_cash: dashboard.cash,
        last_cash_number: record.cashNumber,
        last_dashboard_day: dashboard.day,
        last_dashboard_day_number: dashboardDayNumber(record),
        last_warehouse_inventory: record.warehouseInventory,
        last_warehouse_day: inventoryTable.latestWarehouse.day,
        last_target_team: standingReport.target.team,
        last_target_rank: standingReport.target.rank,
        last_target_cash: standingReport.target.cash,
        last_target_cash_number: standingReport.target.cashNumber,
        last_alerts: warningAlerts,
        last_operational_metrics: operationalSnapshot.metrics,
        last_adjustment_plan: adjustmentPlan,
        last_backtest: {
          generated_at: backtestReport.generated_at,
          data_source: backtestReport.data_source,
          local_dependency: backtestReport.local_dependency,
          observations: backtestReport.observations,
          day_start: backtestReport.day_start,
          day_end: backtestReport.day_end,
          latest: backtestReport.latest,
          recommendations: backtestReport.recommendations,
        },
        last_day_change_review_at: dayChangeReview.changed
          ? record.checkedAt
          : previousState.last_day_change_review_at || null,
        last_day_change_dashboard_day: dayChangeReview.changed
          ? record.dashboardDay
          : previousState.last_day_change_dashboard_day || null,
        last_day_change_review: dayChangeReview,
        last_policy_pages: policySnapshot,
        last_email_sent_at: emailSent
          ? record.checkedAt
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

  console.log(`${warningMinutes}-minute warning email summary:`);
  console.log(`Email sent: ${emailSent ? "yes" : "no"}`);
  console.log(`Target team: ${record.targetTeam}`);
  console.log(`Target cash: ${record.targetCash}`);
  console.log(`Dashboard day: ${record.dashboardDay}`);
  console.log(`Sim day changed: ${dayChangeReview.changed ? "yes" : "no"}`);
  console.log(`Previous dashboard day: ${dayChangeReview.previous_day ?? "n/a"}`);
  console.log(`Remaining game days: ${dayChangeReview.remaining_days ?? "n/a"}`);
  console.log(`Warehouse inventory: ${record.warehouseInventory}`);
  console.log(`Inventory reference crossed: ${record.inventoryCheckpoint ? "yes" : "no"}`);
  console.log(`Active inventory alert: ${record.inventoryAlert ? "yes" : "no"}`);
  console.log(`Warning rule alerts: ${warningAlerts.length}`);
  console.log(`Backtest workbook: ${backtestOutputs.workbookPath || "not written"}`);
  console.log(`Excel data workbook: ${workbookWritten ? dataWorkbookPath : "not written"}`);
  for (const alert of warningAlerts) {
    console.log(
      `- ${alert.label}: ${alert.currentRaw || "n/a"} ${alert.operator} ${alert.thresholdRaw}`,
    );
  }
}

async function runBacktestOnly(config) {
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));
  const { dashboard, inventoryTable, plotSnapshots, standingReport, policySnapshot } =
    await crawl(config);
  const checkedAt = new Date().toISOString();
  const record = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });
  const backtestReport = buildBacktestReport(
    config,
    record,
    standingReport,
    plotSnapshots,
    policySnapshot,
  );
  const outputs = await writeBacktestOutputs(config, backtestReport);

  console.log("Backtest summary:");
  console.log(`Checked at: ${record.checkedAtLocal} (${config.crawl.timezone})`);
  console.log(`Target team: ${record.targetTeam}`);
  console.log(`Dashboard day: ${record.dashboardDay}`);
  console.log(`Observations: ${backtestReport.observations}`);
  console.log(`Window: ${backtestReport.day_start} to ${backtestReport.day_end}`);
  console.log(`Data source: ${backtestReport.data_source}`);
  console.log(`Local dependency: ${backtestReport.local_dependency}`);
  console.log(`Backtest JSON: ${outputs.jsonPath}`);
  console.log(`Backtest summary: ${outputs.summaryCsvPath}`);
  console.log(`Backtest workbook: ${outputs.workbookPath || "not written"}`);
}

async function runPolicyApply(config) {
  ensureDir(path.resolve(process.cwd(), config.output.state_dir));
  const statePath = path.resolve(process.cwd(), config.output.latest_json);
  const previousState = readJson(statePath, {});
  const mode = optionalEnv("POLICY_APPLY_MODE", "recommended").toLowerCase();
  const confirm = optionalEnv("POLICY_APPLY_CONFIRM", "DRY_RUN").trim();
  const autopilotRun = isTruthyEnv("POLICY_APPLY_AUTOPILOT");
  let dryRun = isTruthyEnv("POLICY_APPLY_DRY_RUN") || confirm !== "APPLY";
  const allowShipping =
    isTruthyEnv("POLICY_ALLOW_SHIPPING_METHOD_CHANGE") &&
    config.policy_apply?.allow_shipping_method_change === true;
  const checkedAt = new Date().toISOString();
  const {
    dashboard,
    inventoryTable,
    plotSnapshots,
    standingReport,
    policySnapshot,
    cookieJar,
  } = await crawl(config, { returnCookieJar: true });
  const record = createRecord(config, dashboard, inventoryTable, standingReport, {
    checkedAt,
  });
  const operationalSnapshot = buildOperationalSnapshot(plotSnapshots, {});
  const metricCatalog = buildMetricCatalog(
    config,
    record,
    standingReport,
    operationalSnapshot,
    { policySnapshot },
  );
  const adjustmentPlan = buildAutoAdjustmentPlan(
    config,
    metricCatalog,
    record,
    standingReport,
    policySnapshot,
  );
  const recommendedSet = candidatePolicyChangesFromPlan(
    config,
    adjustmentPlan,
    policySnapshot,
    { allowShipping },
  );
  const candidateSet =
    mode === "custom"
      ? customPolicyChangesFromEnv(config, policySnapshot, recommendedSet)
      : recommendedSet;
  const validated = validatePolicyChanges(config, candidateSet.changes, {
    record,
    metricCatalog,
    policySnapshot,
    allowShipping,
  });
  const accepted = validated.filter((change) => change.accepted);
  const autopilotDecision = buildAutopilotDecision(
    config,
    previousState,
    record,
    validated,
    candidateSet,
    {
      autopilot: autopilotRun,
      dryRun,
    },
  );

  if (autopilotRun && !autopilotDecision.apply_allowed) {
    dryRun = true;
  }

  const report = {
    generated_at: checkedAt,
    generated_at_local: localizedTime(checkedAt, config.crawl.timezone),
    mode,
    dry_run: dryRun,
    confirm,
    workflow_url: policyApplyWorkflowUrl(config),
    target_team: record.targetTeam,
    target_rank: record.targetRank,
    target_cash: record.targetCash,
    dashboard_day: record.dashboardDay,
    warehouse_inventory: record.warehouseInventory,
    warehouse_inventory_day: record.warehouseDay,
    posture: adjustmentPlan.posture || "n/a",
    autopilot: autopilotDecision,
    safety: {
      approval_required: !autopilotRun,
      game_updates_enabled: !dryRun,
      allow_shipping_method_change: allowShipping,
      note: dryRun
        ? autopilotRun
          ? `Autopilot did not submit game forms: ${autopilotDecision.reason}`
          : "Dry run only; no game forms were submitted."
        : autopilotRun
          ? "Autopilot approved this run after latest crawl, consecutive-signal confirmation, and guardrail checks."
          : "Manual approval received; accepted fields were submitted after latest crawl and guardrail checks.",
    },
    conflicts: candidateSet.conflicts,
    changes: validated,
  };

  if (candidateSet.conflicts.length > 0) {
    report.safety.game_updates_enabled = false;
    report.safety.note = "Conflicting recommendations detected; no game forms were submitted.";
  } else if (!dryRun && accepted.length > 0) {
    const byArea = new Map();

    for (const change of accepted) {
      if (!byArea.has(change.area)) {
        byArea.set(change.area, []);
      }
      byArea.get(change.area).push(change);
    }

    for (const [area, changes] of byArea.entries()) {
      await submitPolicyChangesForArea(config, cookieJar, area, changes);
      for (const change of changes) {
        change.applied = true;
      }
    }

    const afterPolicySnapshot = [...policySnapshot];
    for (const area of byArea.keys()) {
      const page = policyPageConfig(config, area);
      const after = await fetchPolicySnapshotForPage(config, cookieJar, page);
      const index = afterPolicySnapshot.findIndex((item) => item.id === page.id);

      if (index >= 0) {
        afterPolicySnapshot[index] = after.snapshot;
      } else {
        afterPolicySnapshot.push(after.snapshot);
      }
    }

    const verified = verifyAppliedChanges(config, afterPolicySnapshot, accepted);
    for (const verifiedChange of verified) {
      const target = report.changes.find(
        (change) =>
          change.area === verifiedChange.area &&
          change.parameter === verifiedChange.parameter,
      );
      if (target) {
        target.verified = verifiedChange.verified;
        target.verified_value = verifiedChange.verified_value;
      }
    }
  }

  const outputs = writePolicyApplyOutputs(config, report);
  const failedVerification = report.changes.filter(
    (change) => change.applied && !change.verified,
  );
  const policySnapshotCsvPath = path.resolve(
    process.cwd(),
    config.output.policy_snapshot_csv ||
      ".monitor-state/policy_snapshot_latest.csv",
  );
  const adjustmentPlanJsonPath = path.resolve(
    process.cwd(),
    config.output.adjustment_plan_json ||
      ".monitor-state/adjustment_plan_latest.json",
  );
  const adjustmentPlanCsvPath = path.resolve(
    process.cwd(),
    config.output.adjustment_plan_csv ||
      ".monitor-state/adjustment_plan_latest.csv",
  );
  fs.writeFileSync(
    policySnapshotCsvPath,
    buildPolicySnapshotCsv(policySnapshot),
    "utf8",
  );
  fs.writeFileSync(
    adjustmentPlanJsonPath,
    `${JSON.stringify(adjustmentPlan, null, 2)}\n`,
    "utf8",
  );
  fs.writeFileSync(
    adjustmentPlanCsvPath,
    buildAdjustmentPlanCsv(adjustmentPlan),
    "utf8",
  );

  const appliedChanges = report.changes.filter((change) => change.applied);
  const currentDayNumber = dashboardDayNumber(record);
  const nextState = {
    ...previousState,
    last_policy_apply_run_at: checkedAt,
    last_policy_apply_run_at_local: report.generated_at_local,
    last_policy_apply_mode: mode,
    last_policy_apply_dry_run: dryRun,
    last_policy_apply_autopilot: autopilotRun,
    last_policy_apply_applied_count: appliedChanges.length,
    last_policy_apply_verified_count: report.changes.filter((change) => change.verified).length,
    last_policy_apply_report: report,
    last_dashboard_day: dashboard.day,
    last_dashboard_day_number: currentDayNumber,
    last_cash: dashboard.cash,
    last_cash_number: record.cashNumber,
    last_warehouse_inventory: record.warehouseInventory,
    last_warehouse_day: inventoryTable.latestWarehouse.day,
    last_target_team: standingReport.target.team,
    last_target_rank: standingReport.target.rank,
    last_target_cash: standingReport.target.cash,
    last_target_cash_number: standingReport.target.cashNumber,
    last_adjustment_plan: adjustmentPlan,
    last_policy_pages: policySnapshot,
  };

  if (autopilotRun) {
    nextState.last_autopilot_run_at = checkedAt;
    nextState.last_autopilot_run_at_local = report.generated_at_local;
    nextState.last_autopilot_enabled = autopilotDecision.env_enabled;
    nextState.last_autopilot_apply_allowed = autopilotDecision.apply_allowed;
    nextState.last_autopilot_reason = autopilotDecision.reason;
    nextState.last_autopilot_signature = autopilotDecision.current_signature;
    nextState.last_autopilot_signature_count =
      autopilotDecision.current_signature
        ? autopilotDecision.consecutive_count
        : 0;
    nextState.last_autopilot_signature_day_number =
      Number.isFinite(currentDayNumber) ? currentDayNumber : null;

    if (appliedChanges.length > 0 && Number.isFinite(currentDayNumber)) {
      nextState.last_autopilot_apply_at = checkedAt;
      nextState.last_autopilot_apply_at_local = report.generated_at_local;
      nextState.last_autopilot_apply_day_number = currentDayNumber;
      nextState.last_autopilot_apply_signature =
        autopilotDecision.current_signature;
    }
  }

  fs.writeFileSync(
    statePath,
    `${JSON.stringify(nextState, null, 2)}\n`,
    "utf8",
  );

  console.log("Policy apply summary:");
  console.log(`Mode: ${mode}`);
  console.log(`Dry run: ${dryRun ? "yes" : "no"}`);
  console.log(`Autopilot: ${autopilotRun ? "yes" : "no"}`);
  if (autopilotRun) {
    console.log(`Autopilot decision: ${autopilotDecision.apply_allowed ? "allow" : "block"} - ${autopilotDecision.reason}`);
    console.log(`Autopilot signature: ${autopilotDecision.current_signature || "n/a"}`);
    console.log(`Autopilot confirmations: ${autopilotDecision.consecutive_count}/${autopilotDecision.consecutive_required}`);
    console.log(`Last autopilot apply day: ${autopilotDecision.last_apply_day ?? "n/a"}`);
  }
  console.log(`Target team: ${record.targetTeam}`);
  console.log(`Target cash: ${record.targetCash}`);
  console.log(`Dashboard day: ${record.dashboardDay}`);
  console.log(`Posture: ${report.posture}`);
  console.log(`Candidate changes: ${candidateSet.changes.length}`);
  console.log(`Accepted changes: ${accepted.length}`);
  console.log(`Conflicts: ${candidateSet.conflicts.length}`);
  for (const change of report.changes) {
    console.log(
      `${change.accepted ? "ACCEPT" : "BLOCK"} ${change.area}.${change.parameter}: ${change.current ?? "n/a"} -> ${change.suggested ?? "n/a"}${change.reject_reason ? ` (${change.reject_reason})` : ""}`,
    );
  }
  console.log(`Policy apply JSON: ${outputs.jsonPath}`);
  console.log(`Policy apply CSV: ${outputs.csvPath}`);

  if (candidateSet.conflicts.length > 0 && !dryRun) {
    throw new Error("Policy apply blocked because recommendation conflicts were detected.");
  }

  if (failedVerification.length > 0) {
    throw new Error(
      `Policy apply submitted but verification failed for: ${failedVerification
        .map((item) => `${item.area}.${item.parameter}`)
        .join(", ")}`,
    );
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

  if (process.argv.includes("--test-email")) {
    await sendTestEmails(config);
    return;
  }

  if (process.argv.includes("--warning-email")) {
    await sendWarningEmail(config);
    return;
  }

  if (process.argv.includes("--backtest")) {
    await runBacktestOnly(config);
    return;
  }

  if (process.argv.includes("--apply-policy")) {
    await runPolicyApply(config);
    return;
  }

  await runOnce(config);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});
