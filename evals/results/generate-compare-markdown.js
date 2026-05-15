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

const isPrimitive = (value) =>
  value === null || ["string", "number", "boolean"].includes(typeof value);

const kvTable = (title, obj) => {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) return "";
  const entries = Object.entries(obj).filter(([, value]) => isPrimitive(value));
  if (entries.length === 0) return "";
  const rows = entries
    .map(([key, value]) => `| ${escapeCell(key)} | ${escapeCell(value)} |`)
    .join("\n");
  return `### ${title}\n\n| Metric | Value |\n| --- | --- |\n${rows}\n`;
};

const arrayTable = (title, arr) => {
  if (!Array.isArray(arr) || arr.length === 0) return "";

  const objectRows = arr.filter(
    (item) => item && typeof item === "object" && !Array.isArray(item),
  );
  if (objectRows.length === 0) {
    const rows = arr
      .slice(0, 20)
      .map(
        (item, index) => `| ${index} | ${escapeCell(JSON.stringify(item))} |`,
      )
      .join("\n");
    return `### ${title}\n\n| Index | Value |\n| --- | --- |\n${rows}\n`;
  }

  const columnSet = new Set();
  objectRows.forEach((row) => {
    Object.entries(row).forEach(([key, value]) => {
      if (isPrimitive(value)) {
        columnSet.add(key);
      }
    });
  });

  let columns = Array.from(columnSet);
  if (columns.length === 0) {
    const rows = objectRows
      .slice(0, 20)
      .map(
        (item, index) => `| ${index} | ${escapeCell(JSON.stringify(item))} |`,
      )
      .join("\n");
    return `### ${title}\n\n| Index | Value |\n| --- | --- |\n${rows}\n`;
  }

  columns = columns.slice(0, 8);
  const header = `| ${columns.map(escapeCell).join(" | ")} |`;
  const separator = `| ${columns.map(() => "---").join(" | ")} |`;
  const rows = objectRows
    .slice(0, 50)
    .map(
      (row) =>
        `| ${columns.map((column) => escapeCell(row[column] ?? "")).join(" | ")} |`,
    )
    .join("\n");

  return `### ${title}\n\n${header}\n${separator}\n${rows}\n`;
};

let md = "## Waza Comparison Report\n\n";
let parsed;
let parseError = "";

if (fs.existsSync(jsonPath)) {
  try {
    parsed = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
  } catch (error) {
    parseError = error instanceof Error ? error.message : String(error);
  }
}

if (parsed !== undefined) {
  if (Array.isArray(parsed)) {
    md +=
      arrayTable("Comparison", parsed) ||
      "No comparable rows were found in compare output.\n";
  } else if (parsed && typeof parsed === "object") {
    const sections = [];
    sections.push(
      kvTable(
        "Summary",
        parsed.summary && typeof parsed.summary === "object"
          ? parsed.summary
          : parsed,
      ),
    );

    const preferredArrayKeys = [
      "models",
      "results",
      "comparisons",
      "items",
      "tasks",
      "scores",
      "rows",
      "data",
    ];
    const seen = new Set();

    preferredArrayKeys.forEach((key) => {
      if (Array.isArray(parsed[key])) {
        sections.push(arrayTable(key, parsed[key]));
        seen.add(key);
      }
    });

    Object.entries(parsed).forEach(([key, value]) => {
      if (!seen.has(key) && Array.isArray(value)) {
        sections.push(arrayTable(key, value));
      }
    });

    md += sections.filter(Boolean).join("\n");
    if (!sections.some(Boolean)) {
      md += "No tabular data found in compare output.\n";
    }
  } else {
    md += "Compare JSON output is not an object or array.\n";
  }
} else {
  md += "Failed to parse compare JSON output.\n";
  if (parseError) {
    md += `\nParse error: ${parseError}\n`;
  }
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
