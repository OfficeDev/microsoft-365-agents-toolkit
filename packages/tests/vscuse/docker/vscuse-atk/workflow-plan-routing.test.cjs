const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const { createServer } = require("node:http");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const { parse } = require("yaml");

const workflowsDirectory = path.resolve(
  __dirname,
  "../../../../../.github/workflows",
);

function readWorkflow(fileName) {
  return parse(
    fs.readFileSync(path.join(workflowsDirectory, fileName), "utf8"),
  );
}

test("VTR-01: template workflow excludes separately scheduled plans", () => {
  const workflow = readWorkflow("ui-test-vscuse-template.yml");

  assert.equal(
    workflow.jobs.run.with.plan_find_args,
    '-name "*.json" ! -name "Feature_*" ! -name "feature-*" ! -name "Sample_*" ! -name "Regular_*"',
  );
});

test("VTR-04: feature workflow selects legacy and generated feature plans", () => {
  const workflow = readWorkflow("ui-test-vscuse-features.yml");

  assert.equal(
    workflow.jobs.run.with.plan_find_args,
    '\\( -name "Feature_*.json" -o -name "feature-*.json" \\)',
  );
});

test("VTR-05: PR selector recognizes generated feature plan names", () => {
  const workflow = readWorkflow("pr-vscuse-test.yml");
  const selectorStep = workflow.jobs["select-plans-and-run"].steps.find(
    (step) => step.id === "ai-select",
  );

  assert.ok(selectorStep);
  assert.match(selectorStep.run, /GENERATED FEATURE PLAN RULE/);
  assert.match(selectorStep.run, /feature-\*/);
  assert.match(selectorStep.run, /feature-da-\*/);
  assert.match(selectorStep.run, /do not require.*_ts.*_js.*_py/i);

  const basePromptIndex = selectorStep.run.indexOf(
    "SYSTEM_PROMPT='You are a test plan selector",
  );
  const generatedFeatureRuleIndex = selectorStep.run.indexOf(
    "GENERATED FEATURE PLAN RULE",
  );
  const userPromptIndex = selectorStep.run.indexOf("USER_PROMPT=$(printf");

  assert.ok(basePromptIndex >= 0);
  assert.ok(generatedFeatureRuleIndex > basePromptIndex);
  assert.ok(userPromptIndex > generatedFeatureRuleIndex);
});

test("VTR-06: shared workflow supplies retained feature auth fixtures", () => {
  const workflow = readWorkflow("ui-test-vscuse-common.yml");
  const environment = workflow.jobs.main.env;

  assert.equal(
    environment.EXISTING_ENTRA_CLIENT_ID,
    "${{ secrets.EXISTING_ENTRA_CLIENT_ID || '4cfde729-32e4-4862-a409-07e14dbfd296' }}",
  );
  assert.equal(
    environment.EXISTING_GITHUB_OAUTH_CLIENT_ID,
    "${{ secrets.EXISTING_GITHUB_OAUTH_CLIENT_ID || 'fakefakefakefakeid1' }}",
  );
  assert.equal(
    environment.EXISTING_GITHUB_OAUTH_CLIENT_SECRET,
    "${{ secrets.EXISTING_GITHUB_OAUTH_CLIENT_SECRET || 'fakefakefakefakefakefakefakefakefakefake' }}",
  );
});

test("VTR-02: regular workflow supports manual and weekly runs", () => {
  const workflow = readWorkflow("ui-test-vscuse-regular.yml");

  assert.ok(workflow.on.workflow_dispatch);
  assert.deepEqual(workflow.on.schedule, [{ cron: "0 2 * * 1" }]);
});

test("VTR-03: regular workflow selects Regular plans through the shared workflow", () => {
  const workflow = readWorkflow("ui-test-vscuse-regular.yml");

  assert.equal(
    workflow.jobs.run.uses,
    "./.github/workflows/ui-test-vscuse-common.yml",
  );
  assert.equal(workflow.jobs.run.with.plan_find_args, '-name "Regular_*.json"');
  assert.equal(
    workflow.jobs.run.with.schedule_trigger,
    "${{ github.event_name == 'schedule' || github.event.inputs.schedule_trigger == 'true' }}",
  );
  assert.equal(workflow.jobs.run.with.email_subject_suffix, "[Regular]");
});

test("VTR-07: all vscuse callers use APIM reporting without an authentication switch", () => {
  const shared = readWorkflow("ui-test-vscuse-common.yml");
  for (const group of ["template", "features", "regular", "samples"]) {
    const caller = readWorkflow(`ui-test-vscuse-${group}.yml`);
    assert.equal(
      caller.jobs.run.uses,
      "./.github/workflows/ui-test-vscuse-common.yml",
    );
    assert.doesNotMatch(JSON.stringify(caller), /\buse_apim\b/i);
  }
  assert.doesNotMatch(JSON.stringify(shared), /\buse_apim\b/i);
  assert.doesNotMatch(JSON.stringify(shared), /az storage blob upload/);
  assert.ok(
    shared.jobs.main.steps.every(
      (step) => !step.uses?.startsWith("azure/login@"),
    ),
  );
  assert.ok(
    shared.jobs.report.steps.some((step) =>
      step.uses?.startsWith("azure/login@"),
    ),
  );
});

test("VTR-10: APIM uses a pinned runner with distinct keys and preserves test-case config", (context) => {
  const workflow = readWorkflow("ui-test-vscuse-common.yml");
  const steps = workflow.jobs.main.steps;
  const download = steps.find((step) => step.name === "Run download script");
  const configure = steps.find(
    (step) => step.name === "Configure vscuse APIM clients",
  );
  const run = steps.find((step) => step.name === "Run test");

  assert.match(download.run, /node download-vscuse-python\.js v0\.2\.67\s/);
  assert.equal(workflow.on.workflow_call.inputs.vscuse_version, undefined);
  for (const group of ["template", "features", "regular", "samples"]) {
    const caller = readWorkflow(`ui-test-vscuse-${group}.yml`);
    assert.doesNotMatch(JSON.stringify(caller), /\bvscuse_version\b/i);
  }
  assert.equal(configure.if, undefined);
  assert.equal(
    workflow.jobs.main.env.VSCUSE_AZURE_OPENAI_ENDPOINT,
    "https://apim-ai-atktest1022.azure-api.net",
  );
  assert.equal(
    workflow.jobs.main.env.VSCUSE_AZURE_VISION_ENDPOINT,
    "https://apim-ai-atktest1022.azure-api.net/vscuse-vision",
  );
  assert.match(
    run.env.VSCUSE_FOUNDRY_KEY,
    /secrets\.VSCUSE_TEST_APIM_AZURE_OPENAI_KEY/,
  );
  assert.match(run.env.VSCUSE_VISION_KEY, /secrets\.VSCUSE_VISION_APIM_KEY/);

  const directory = reportFixture(context);
  const source = path.resolve(
    workflowsDirectory,
    "../../packages/tests/vscuse/vscode-test-cases/config.yaml",
  );
  const config = path.join(directory, "config.yaml");
  fs.copyFileSync(source, config);
  const before = parse(fs.readFileSync(config, "utf8"));
  const node = configure.run.match(/node <<'NODE'\n([\s\S]+?)\nNODE/)[1];
  const result = spawnSync(process.execPath, ["-"], {
    cwd: directory,
    input: node,
    encoding: "utf8",
    env: {
      ...process.env,
      NODE_PATH: path.dirname(
        path.dirname(require.resolve("yaml/package.json")),
      ),
    },
  });
  assert.equal(result.status, 0, result.stderr);
  const after = parse(fs.readFileSync(config, "utf8"));
  assert.deepEqual(after.docker, before.docker);
  assert.deepEqual(after.azure_openai, {
    ...before.azure_openai,
    use_azure_credential: false,
    api_key: "${VSCUSE_FOUNDRY_KEY}",
  });
  assert.deepEqual(after.azure_vision, {
    ...before.azure_vision,
    use_azure_credential: false,
    api_key: "${VSCUSE_VISION_KEY}",
  });
});

test("VTR-11: failed Azure DevOps login remains a failure without blocking APIM summary upload", () => {
  const steps = readWorkflow("ui-test-vscuse-common.yml").jobs.report.steps;
  const login = steps.find((step) => step.id === "report-azure-login");
  const listJobs = steps.find((step) => step.id === "list-jobs");
  const upload = steps.find(
    (step) => step.name === "Upload aggregated report to Azure Blob Storage",
  );
  assert.equal(login["continue-on-error"], undefined);
  assert.ok(
    steps.findIndex((step) => step.name === "Download storage URLs") <
      steps.indexOf(login),
  );
  assert.match(listJobs.if, /steps\.report-azure-login\.outcome == 'failure'/);
  assert.match(listJobs.run, /Skipped Azure DevOps bug enrichment/);
  assert.match(upload.if, /steps\.list-jobs\.outcome == 'success'/);
});

async function captureUploads(context, status = 201) {
  const uploads = [];
  const server = createServer(async (request, response) => {
    const buffers = [];
    for await (const chunk of request) {
      buffers.push(chunk);
    }
    uploads.push({
      method: request.method,
      url: request.url,
      key: request.headers["ocp-apim-subscription-key"],
      length: Number(request.headers["content-length"]),
      type: request.headers["content-type"],
      body: Buffer.concat(buffers),
    });
    response.writeHead(status);
    response.end();
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(() => new Promise((resolve) => server.close(resolve)));
  return {
    uploads,
    endpoint: `http://127.0.0.1:${server.address().port}/vscuse-report-upload`,
  };
}

function reportFixture(context) {
  const directory = fs.mkdtempSync(
    path.join(os.tmpdir(), "vscuse-upload-test-"),
  );
  context.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  return directory;
}

test("VTR-08: reports stream their original bytes to APIM and preserve report URLs", async (context) => {
  const {
    uploadReport,
  } = require("../../../../../.github/scripts/upload-vscuse-report.cjs");
  const directory = reportFixture(context);
  const casePath = path.join(directory, "case");
  const assetPath = path.join(casePath, "assets");
  const summaryPath = path.join(directory, "email-body.html");
  const asset = Buffer.from([0, 1, 255]);
  fs.mkdirSync(assetPath, { recursive: true });
  fs.writeFileSync(path.join(casePath, "index.html"), "<p>case</p>");
  fs.writeFileSync(path.join(assetPath, "test image.png"), asset);
  fs.writeFileSync(summaryPath, "<p>summary</p>");

  const { endpoint, uploads } = await captureUploads(context);
  const guid = "16efbb47-18da-4baa-8493-20781ef8d664";
  const options = { endpoint, key: "test-key" };
  await uploadReport(casePath, guid, options);
  await uploadReport(summaryPath, guid, options);
  const upload = (filePath, type, body) => ({
    method: "PUT",
    url: `/vscuse-report-upload/${guid}/${filePath}`,
    key: "test-key",
    length: body.length,
    type,
    body,
  });
  const expectedCase = [
    upload("index.html", "text/html", Buffer.from("<p>case</p>")),
    upload("assets/test%20image.png", "image/png", asset),
  ];
  assert.deepEqual(
    uploads.slice(0, 2).sort((a, b) => a.url.localeCompare(b.url)),
    expectedCase.sort((a, b) => a.url.localeCompare(b.url)),
  );
  assert.deepEqual(uploads.slice(2), [
    upload("index.html", "text/html", Buffer.from("<p>summary</p>")),
  ]);

  const shared = readWorkflow("ui-test-vscuse-common.yml");
  for (const [job, name, marker] of [
    ["main", "🌐Upload files to Azure Blob Storage", "Here is the report:"],
    [
      "report",
      "Upload aggregated report to Azure Blob Storage",
      "Report Summary URL:",
    ],
  ]) {
    const script = shared.jobs[job].steps.find(
      (step) => step.name === name,
    ).run;
    assert.match(
      script,
      /storproxy-app-voazuxhhvtgiq\.azurewebsites\.net\/\$guid\/index\.html/,
    );
    assert.match(
      script,
      /upload-vscuse-report\.cjs[^\n]*\n\s*if \(\$LASTEXITCODE -ne 0\) \{ exit \$LASTEXITCODE \}/,
    );
    assert.ok(
      script.indexOf(marker) > script.indexOf("upload-vscuse-report.cjs"),
    );
  }
});

test("VTR-09: APIM uploads fail for missing inputs, oversized files, or a non-201 response", async (context) => {
  const {
    uploadReport,
  } = require("../../../../../.github/scripts/upload-vscuse-report.cjs");
  const directory = reportFixture(context);
  const { endpoint, uploads } = await captureUploads(context, 200);
  const guid = "98c559dd-db07-4464-9d1a-ef8bd11d7550";
  const options = { endpoint, key: "test-key" };
  await assert.rejects(
    uploadReport(directory, guid, { endpoint, key: "" }),
    /VSCUSE_REPORT_APIM_KEY/,
  );
  await assert.rejects(uploadReport(directory, guid, options), /index\.html/);
  const indexPath = path.join(directory, "index.html");
  fs.writeFileSync(indexPath, "");
  fs.truncateSync(indexPath, 100 * 1024 * 1024 + 1);
  await assert.rejects(uploadReport(directory, guid, options), /100 MiB/);
  assert.equal(uploads.length, 0);

  fs.writeFileSync(indexPath, "<p>report</p>");
  await assert.rejects(uploadReport(directory, guid, options), /HTTP 200/);
  assert.equal(uploads.length, 1);
});
