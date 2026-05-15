const fs = require("fs");
const core = require("@actions/core");
const github = require("@actions/github");

async function run() {
  try {
    const marker = "<!-- waza-compare-report -->";
    const reportPath = "evals/results/compare.md";
    let report = "Comparison report was not generated.";
    if (fs.existsSync(reportPath)) {
      report = fs.readFileSync(reportPath, "utf8");
    }
    if (report.length > 60000) {
      report = `${report.slice(0, 60000)}\n\n... report truncated ...`;
    }
    const body = [
      marker,
      "## Waza Compare Result",
      "",
      "Baseline: `evals/results/baseline.json`",
      "Latest: `evals/results/latest.json`",
      "",
      report,
    ].join("\n");

    const token =
      process.env.GITHUB_TOKEN ||
      process.env.GITHUB_TOKEN ||
      core.getInput("github-token");
    if (!token) {
      throw new Error("GITHUB_TOKEN is required");
    }
    const octokit = github.getOctokit(token);
    const { owner, repo } = github.context.repo;
    const issue_number = github.context.payload.pull_request.number;
    const comments = await octokit.paginate(octokit.rest.issues.listComments, {
      owner,
      repo,
      issue_number,
      per_page: 100,
    });
    const existing = comments.find(
      (comment) =>
        comment.user?.type === "Bot" && comment.body?.includes(marker),
    );
    if (existing) {
      await octokit.rest.issues.updateComment({
        owner,
        repo,
        comment_id: existing.id,
        body,
      });
    } else {
      await octokit.rest.issues.createComment({
        owner,
        repo,
        issue_number,
        body,
      });
    }
  } catch (error) {
    core.setFailed(error.message);
  }
}

run();
