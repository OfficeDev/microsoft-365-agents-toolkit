const fs = require("fs");

const comparePath = "evals/results/compare.json";
const baselinePath = "evals/results/baseline.json";
const latestPath = "evals/results/latest.json";
const outputPath = "evals/results/compare-dashboard.html";

const readJson = (filePath) => {
  if (!fs.existsSync(filePath)) {
    return null;
  }

  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return null;
  }
};

const toNumber = (value) =>
  typeof value === "number" && Number.isFinite(value) ? value : null;
const fmtInt = (value) => {
  const n = toNumber(value);
  return n === null ? "n/a" : Math.round(n).toLocaleString("en-US");
};
const pct = (value) => {
  const n = toNumber(value);
  return n === null ? "n/a" : `${(n * 100).toFixed(1)}%`;
};
const num = (value, digits = 3) => {
  const n = toNumber(value);
  return n === null ? "n/a" : n.toFixed(digits);
};
const ms = (value) => {
  const n = toNumber(value);
  return n === null ? "n/a" : `${Math.round(n)} ms`;
};
const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");

const totalTokensFromUsage = (usage) => {
  if (!usage || typeof usage !== "object") {
    return null;
  }

  const keys = [
    "input_tokens",
    "output_tokens",
    "cache_read_tokens",
    "cache_write_tokens",
  ];
  let total = 0;
  let hasValue = false;

  for (const key of keys) {
    const n = toNumber(usage[key]);
    if (n !== null) {
      total += n;
      hasValue = true;
    }
  }

  return hasValue ? total : null;
};

const extractTaskMetrics = (result) => {
  const map = new Map();
  const tasks = Array.isArray(result?.tasks) ? result.tasks : [];

  for (const task of tasks) {
    const runs = Array.isArray(task?.runs) ? task.runs : [];

    const durationMs = runs.reduce(
      (sum, run) => sum + (toNumber(run?.duration_ms) ?? 0),
      0,
    );
    const tokenCount = runs.reduce((sum, run) => {
      const v1 = toNumber(
        run?.validations?.efficiency_check?.details?.tokens_total,
      );
      const v2 = toNumber(run?.session_digest?.token_usage?.total_tokens);
      const v3 = toNumber(run?.usage?.total_tokens);
      return sum + (v1 ?? v2 ?? v3 ?? 0);
    }, 0);

    map.set(task?.test_id, {
      taskId: task?.test_id,
      displayName: task?.display_name ?? task?.test_id ?? "unknown",
      durationMs,
      tokenCount,
    });
  }

  return map;
};

let compare = null;
let parseError = null;

try {
  compare = readJson(comparePath);
  if (!compare) {
    parseError = "compare.json is missing or invalid.";
  }
} catch (error) {
  parseError = error instanceof Error ? error.message : String(error);
}

const baselineResult = readJson(baselinePath);
const latestResult = readJson(latestPath);

const files = Array.isArray(compare?.files) ? compare.files : [];
const aggregateScores = Array.isArray(compare?.aggregate_scores)
  ? compare.aggregate_scores
  : [];
const successRates = Array.isArray(compare?.success_rates)
  ? compare.success_rates
  : [];
const durations = Array.isArray(compare?.durations_ms)
  ? compare.durations_ms
  : [];
const taskDeltas = Array.isArray(compare?.task_deltas)
  ? compare.task_deltas
  : [];

const baselineUsage = baselineResult?.summary?.usage ?? null;
const latestUsage = latestResult?.summary?.usage ?? null;

const baselineTotalTokens = totalTokensFromUsage(baselineUsage);
const latestTotalTokens = totalTokensFromUsage(latestUsage);
const totalTokenDelta =
  baselineTotalTokens !== null && latestTotalTokens !== null
    ? latestTotalTokens - baselineTotalTokens
    : null;

const baselineTaskMetrics = extractTaskMetrics(baselineResult);
const latestTaskMetrics = extractTaskMetrics(latestResult);

const pctDelta = (base, delta) => {
  const b = toNumber(base);
  const d = toNumber(delta);
  if (b === null || d === null || b === 0) return "n/a";
  const p = (d / Math.abs(b)) * 100;
  const sign = p > 0 ? "+" : "";
  return `${sign}${p.toFixed(1)}%`;
};

const deltaClass = (value) => {
  const n = toNumber(value);
  if (n === null || n === 0) return "neutral";
  return n > 0 ? "positive" : "negative";
};

const baselineLabel = escapeHtml(files[0] ?? "baseline");
const latestLabel = escapeHtml(files[1] ?? "latest");

const summaryCards = `
  <section class="grid cards">
    <article class="card">
      <h3>Aggregate Score</h3>
      <div class="split"><span>${num(aggregateScores[0])}</span><span>${num(aggregateScores[1])}</span></div>
      <p class="delta ${deltaClass(compare?.aggregate_score_delta)}">Delta: ${num(compare?.aggregate_score_delta, 4)}</p>
    </article>
    <article class="card">
      <h3>Success Rate</h3>
      <div class="split"><span>${pct(successRates[0])}</span><span>${pct(successRates[1])}</span></div>
      <p class="delta ${deltaClass(compare?.success_rate_delta)}">Delta: ${pct(compare?.success_rate_delta)}</p>
    </article>
    <article class="card">
      <h3>Duration</h3>
      <div class="split"><span>${ms(durations[0])}</span><span>${ms(durations[1])}</span></div>
      <p class="delta ${deltaClass(compare?.duration_delta_ms)}">Delta: ${pctDelta(durations[0], compare?.duration_delta_ms)}</p>
    </article>
    <article class="card">
      <h3>Total Token Usage</h3>
      <div class="split"><span>${fmtInt(baselineTotalTokens)}</span><span>${fmtInt(latestTotalTokens)}</span></div>
      <p class="delta ${deltaClass(totalTokenDelta)}">Delta: ${pctDelta(baselineTotalTokens, totalTokenDelta)}</p>
    </article>
  </section>
`;

const maxScoreDelta = Math.max(
  0.001,
  ...taskDeltas.map((item) => Math.abs(toNumber(item?.score_delta) ?? 0)),
);

const taskRows = taskDeltas.length
  ? taskDeltas
      .map((item) => {
        const scoreDelta = toNumber(item?.score_delta) ?? 0;
        const passRateDelta = toNumber(item?.pass_rate_delta) ?? 0;
        const width = `${Math.max(4, Math.round((Math.abs(scoreDelta) / maxScoreDelta) * 100))}%`;
        const taskId = item?.task_id;
        const baselineTask = baselineTaskMetrics.get(taskId) || {};
        const latestTask = latestTaskMetrics.get(taskId) || {};
        const durationDelta =
          typeof latestTask.durationMs === "number" &&
          typeof baselineTask.durationMs === "number"
            ? latestTask.durationMs - baselineTask.durationMs
            : null;
        const tokenDelta =
          typeof latestTask.tokenCount === "number" &&
          typeof baselineTask.tokenCount === "number"
            ? latestTask.tokenCount - baselineTask.tokenCount
            : null;

        return `
          <tr>
            <td>${escapeHtml(item?.display_name ?? item?.task_id ?? "unknown")}</td>
            <td class="mono">${escapeHtml(item?.task_id ?? "n/a")}</td>
            <td>
              <div class="bar-track">
                <div class="bar ${scoreDelta >= 0 ? "bar-pos" : "bar-neg"}" style="width:${width}"></div>
              </div>
              <span class="delta ${deltaClass(scoreDelta)}">${num(scoreDelta, 4)}</span>
            </td>
            <td class="delta ${deltaClass(passRateDelta)}">${pct(passRateDelta)}</td>
          </tr>
          <tr>
            <td colspan="2" class="mono" style="color:var(--muted);font-size:12px;">Duration (ms)</td>
            <td colspan="2">
              <span>${fmtInt(baselineTask.durationMs)}</span> → <span>${fmtInt(latestTask.durationMs)}</span>
              <span class="delta ${deltaClass(durationDelta)}" style="margin-left:12px;">Δ ${pctDelta(baselineTask.durationMs, durationDelta)}</span>
            </td>
          </tr>
          <tr>
            <td colspan="2" class="mono" style="color:var(--muted);font-size:12px;">Token Usage</td>
            <td colspan="2">
              <span>${fmtInt(baselineTask.tokenCount)}</span> → <span>${fmtInt(latestTask.tokenCount)}</span>
              <span class="delta ${deltaClass(tokenDelta)}" style="margin-left:12px;">Δ ${pctDelta(baselineTask.tokenCount, tokenDelta)}</span>
            </td>
          </tr>
        `;
      })
      .join("\n")
  : '<tr><td colspan="4">No task delta data found.</td></tr>';

const body = parseError
  ? `<section class=\"card\"><h3>Compare JSON Parse Failed</h3><p>${escapeHtml(parseError)}</p></section>`
  : compare
    ? `
      ${summaryCards}
      <section class="card table-card">
        <h3>Task Delta Details</h3>
        <table>
          <thead>
            <tr>
              <th>Task</th>
              <th>Task ID</th>
              <th>Score Delta</th>
              <th>Pass Rate Delta</th>
            </tr>
          </thead>
          <tbody>
            ${taskRows}
          </tbody>
        </table>
      </section>
    `
    : '<section class="card"><h3>No compare.json found</h3><p>Run compare step first to generate visualization input.</p></section>';

const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Waza Compare Dashboard</title>
  <style>
    :root {
      --bg: #0b1020;
      --panel: #151b31;
      --panel-2: #1a2240;
      --text: #e9eefc;
      --muted: #96a3c8;
      --pos: #39d98a;
      --neg: #ff6b6b;
      --neutral: #aab4d6;
      --accent: #5bd0ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", "Helvetica Neue", sans-serif;
      background: radial-gradient(circle at 20% -10%, #1d2a56, var(--bg) 45%), linear-gradient(180deg, #0f1630, var(--bg));
      color: var(--text);
      min-height: 100vh;
      padding: 24px;
    }
    .wrap { max-width: 1200px; margin: 0 auto; }
    h1 { margin: 0 0 8px; font-size: 30px; letter-spacing: 0.2px; }
    .sub { margin: 0 0 20px; color: var(--muted); font-size: 14px; }
    .grid { display: grid; gap: 16px; }
    .cards { grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); margin-bottom: 16px; }
    .card {
      background: linear-gradient(165deg, var(--panel), var(--panel-2));
      border: 1px solid #2a3768;
      border-radius: 16px;
      padding: 16px;
      box-shadow: 0 14px 30px rgba(0, 0, 0, 0.22);
    }
    .table-card { overflow-x: auto; }
    h3 { margin: 0 0 10px; font-size: 15px; color: var(--accent); }
    .split { display: flex; justify-content: space-between; font-weight: 700; font-size: 18px; margin-bottom: 8px; }
    .delta { font-weight: 700; }
    .delta.positive { color: var(--pos); }
    .delta.negative { color: var(--neg); }
    .delta.neutral { color: var(--neutral); }
    .mono { font-family: Consolas, "Courier New", monospace; font-size: 12px; color: var(--muted); }
    table { width: 100%; border-collapse: collapse; min-width: 860px; }
    th, td { text-align: left; padding: 10px 8px; border-bottom: 1px solid #2b3868; vertical-align: middle; }
    th { color: #b7c4eb; font-size: 12px; text-transform: uppercase; letter-spacing: 0.6px; }
    .bar-track { width: 100%; height: 8px; background: #263056; border-radius: 999px; overflow: hidden; margin: 0 0 6px; }
    .bar { height: 100%; border-radius: 999px; }
    .bar-pos { background: linear-gradient(90deg, #2bc97f, #8cffc7); }
    .bar-neg { background: linear-gradient(90deg, #ff6b6b, #ff9a9a); }
    @media (max-width: 720px) {
      body { padding: 14px; }
      h1 { font-size: 24px; }
    }
  </style>
</head>
<body>
  <main class="wrap">
    <h1>Waza Compare Dashboard</h1>
    <p class="sub">Baseline: ${baselineLabel} | Latest: ${latestLabel}</p>
    ${body}
  </main>
</body>
</html>`;

fs.writeFileSync(outputPath, html, "utf8");
