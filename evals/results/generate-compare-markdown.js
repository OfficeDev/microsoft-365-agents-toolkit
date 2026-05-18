const fs = require("fs");

const jsonPath = "evals/results/compare.json";
const stderrPath = "evals/results/compare.stderr";
const mdPath = "evals/results/compare.md";
const compareExitCode = Number(process.env.COMPARE_EXIT_CODE || "0");

const escapeCell = (value) =>
  String(value)
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\r?\n/g, " ")
    .trim();

const formatFloat = (num) => {
  if (typeof num !== "number") return String(num);
  return num.toFixed(4);
};

const formatDelta = (delta) => {
  if (typeof delta !== "number") return String(delta);
  const formatted = delta.toFixed(4);
  const sign = delta > 0 ? "+" : "";
  return `${sign}${formatted}`;
};

let md = "## Evaluation Comparison Report\n\n";
let parsed;
let parseError = "";

if (fs.existsSync(jsonPath)) {
  try {
    parsed = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
  } catch (error) {
    parseError = error instanceof Error ? error.message : String(error);
  }
}

if (parsed !== undefined && parsed && typeof parsed === "object") {
  // Summary table with Baseline and Latest columns
  const summaryRows = [];

  if (
    Array.isArray(parsed.aggregate_scores) &&
    parsed.aggregate_scores.length >= 2
  ) {
    summaryRows.push(
      `| Aggregate Score | ${formatFloat(parsed.aggregate_scores[0])} | ${formatFloat(parsed.aggregate_scores[1])} | ${formatDelta(parsed.aggregate_score_delta)} |`,
    );
  }

  if (Array.isArray(parsed.success_rates) && parsed.success_rates.length >= 2) {
    summaryRows.push(
      `| Success Rate | ${formatFloat(parsed.success_rates[0])} | ${formatFloat(parsed.success_rates[1])} | ${formatDelta(parsed.success_rate_delta)} |`,
    );
  }

  if (Array.isArray(parsed.durations_ms) && parsed.durations_ms.length >= 2) {
    const durationDelta =
      parsed.duration_delta_ms !== undefined
        ? parsed.duration_delta_ms
        : parsed.durations_ms[1] - parsed.durations_ms[0];
    summaryRows.push(
      `| Duration (ms) | ${parsed.durations_ms[0]} | ${parsed.durations_ms[1]} | ${formatDelta(durationDelta)} |`,
    );
  }

  if (summaryRows.length > 0) {
    md += "### Summary Metrics\n\n";
    md += "| Metric | Baseline | Latest | Delta |\n";
    md += "| --- | --- | --- | --- |\n";
    md += summaryRows.join("\n");
    md += "\n\n";
  }

  // Task Deltas table
  if (Array.isArray(parsed.task_deltas) && parsed.task_deltas.length > 0) {
    md += "### Task Results\n\n";
    md +=
      "| Task | Score (Baseline → Latest) | Pass Rate (Baseline → Latest) | Status |\n";
    md += "| --- | --- | --- | --- |\n";

    parsed.task_deltas.forEach((task) => {
      const taskName = escapeCell(
        task.display_name || task.task_id || "Unknown",
      );
      const baselineScore = formatFloat(task.scores?.[0] ?? "—");
      const latestScore = formatFloat(task.scores?.[1] ?? "—");
      const scoreStr = `${baselineScore} → ${latestScore}`;

      const baselinePassRate = formatFloat(task.pass_rates?.[0] ?? "—");
      const latestPassRate = formatFloat(task.pass_rates?.[1] ?? "—");
      const passRateStr = `${baselinePassRate} → ${latestPassRate}`;

      const baselineStatus = task.statuses?.[0] ?? "—";
      const latestStatus = task.statuses?.[1] ?? "—";
      const statusStr =
        baselineStatus === latestStatus
          ? `✓ ${baselineStatus}`
          : `${baselineStatus} → ${latestStatus}`;

      md += `| ${taskName} | ${scoreStr} | ${passRateStr} | ${statusStr} |\n`;
    });
    md += "\n";
  }

  if (
    summaryRows.length === 0 &&
    (!Array.isArray(parsed.task_deltas) || parsed.task_deltas.length === 0)
  ) {
    md += "No comparison data found.\n";
  }
} else if (parsed === undefined) {
  md += "Failed to parse compare JSON output.\n";
  if (parseError) {
    md += `\nParse error: ${parseError}\n`;
  }
} else {
  md += "Compare JSON output is not an object.\n";
}

if (compareExitCode !== 0) {
  md += `\nComparison command exited with code ${compareExitCode}.\n`;
}

if (fs.existsSync(stderrPath)) {
  const stderr = fs.readFileSync(stderrPath, "utf8").trim();
  if (stderr) {
    md += "\n### Command Stderr\n\n";
    md += "```text\n";
    md += stderr.slice(0, 20000);
    md += "\n```\n";
  }
}

fs.writeFileSync(mdPath, md, "utf8");
