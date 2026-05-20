const fs = require("fs");

const baselinePath = "evals/results/baseline.json";
const latestPath = "evals/results/latest.json";
const outputPath = "evals/results/compare.json";

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

const getSummaryMetric = (result, key) => toNumber(result?.summary?.[key]);

const getTaskStats = (task) => {
  const score = toNumber(task?.stats?.avg_score);
  const passRate = toNumber(task?.stats?.pass_rate);
  const status = typeof task?.status === "string" ? task.status : null;

  return {
    score,
    passRate,
    status,
  };
};

const buildTaskMap = (result) => {
  const map = new Map();
  const tasks = Array.isArray(result?.tasks) ? result.tasks : [];

  tasks.forEach((task) => {
    const id = task?.test_id;
    if (!id) {
      return;
    }

    map.set(id, {
      task_id: id,
      display_name: task?.display_name ?? id,
      ...getTaskStats(task),
    });
  });

  return map;
};

const baseline = readJson(baselinePath);
const latest = readJson(latestPath);

if (!baseline || !latest) {
  console.error(
    "baseline.json or latest.json is missing/invalid, skip compare.json generation.",
  );
  process.exitCode = 1;
} else {
  const baselineTasks = buildTaskMap(baseline);
  const latestTasks = buildTaskMap(latest);
  const allTaskIds = Array.from(
    new Set([...baselineTasks.keys(), ...latestTasks.keys()]),
  );

  const taskDeltas = allTaskIds.map((taskId) => {
    const base = baselineTasks.get(taskId) || {
      task_id: taskId,
      display_name: taskId,
      score: null,
      passRate: null,
      status: null,
    };
    const next = latestTasks.get(taskId) || {
      task_id: taskId,
      display_name: base.display_name,
      score: null,
      passRate: null,
      status: null,
    };

    const scoreDelta =
      base.score !== null && next.score !== null
        ? next.score - base.score
        : null;
    const passRateDelta =
      base.passRate !== null && next.passRate !== null
        ? next.passRate - base.passRate
        : null;

    return {
      task_id: taskId,
      display_name: next.display_name || base.display_name || taskId,
      scores: [base.score, next.score],
      pass_rates: [base.passRate, next.passRate],
      statuses: [base.status, next.status],
      score_delta: scoreDelta,
      pass_rate_delta: passRateDelta,
    };
  });

  const aggregateScores = [
    getSummaryMetric(baseline, "aggregate_score"),
    getSummaryMetric(latest, "aggregate_score"),
  ];
  const successRates = [
    getSummaryMetric(baseline, "success_rate"),
    getSummaryMetric(latest, "success_rate"),
  ];
  const durations = [
    getSummaryMetric(baseline, "duration_ms"),
    getSummaryMetric(latest, "duration_ms"),
  ];
  const totalTests = [
    getSummaryMetric(baseline, "total_tests"),
    getSummaryMetric(latest, "total_tests"),
  ];

  const compare = {
    files: [baselinePath, latestPath],
    models: [
      baseline?.config?.model_id ?? null,
      latest?.config?.model_id ?? null,
    ],
    aggregate_scores: aggregateScores,
    success_rates: successRates,
    aggregate_score_delta:
      aggregateScores[0] !== null && aggregateScores[1] !== null
        ? aggregateScores[1] - aggregateScores[0]
        : null,
    success_rate_delta:
      successRates[0] !== null && successRates[1] !== null
        ? successRates[1] - successRates[0]
        : null,
    task_deltas: taskDeltas,
    total_tests: totalTests,
    durations_ms: durations,
    duration_delta_ms:
      durations[0] !== null && durations[1] !== null
        ? durations[1] - durations[0]
        : null,
  };

  fs.writeFileSync(outputPath, JSON.stringify(compare, null, 2), "utf8");
}
