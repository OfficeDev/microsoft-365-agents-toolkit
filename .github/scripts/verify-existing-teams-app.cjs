const { chromium } = require("playwright");
const { randomUUID } = require("node:crypto");

const appName = "vscuse_app_6Hwdxdev";
const teamsHosts = new Set(["teams.microsoft.com", "teams.cloud.microsoft"]);

async function reportBrowserState(page) {
  const url = new URL(page.url());
  const activeHost = teamsHosts.has(url.hostname)
    ? url.hostname
    : url.hostname === "login.microsoftonline.com"
      ? url.hostname
      : "other";
  const state = await page.evaluate(() => ({
    appsButton: !!document.querySelector('[aria-label="Apps"]'),
    chats: /\bChat\b/.test(document.body?.innerText ?? ""),
    store: /\bManage your apps\b/.test(document.body?.innerText ?? ""),
    loading: /\b(?:Loading|Getting things ready)\b/i.test(document.body?.innerText ?? ""),
    loginForm: !!document.querySelector('input[name="loginfmt"], input[name="passwd"]'),
    docReady: document.readyState,
  }));
  console.log(`Teams browser diagnostics: ${JSON.stringify({ activeHost, ...state })}`);
}

async function run() {
  const users = process.env.M365_USERNAMES.split(",").map((user) => user.trim());
  const user = users[14];
  const password = process.env.M365_TEST_PASSWORD;
  if (users.length !== 19 || !user || !password) {
    throw new Error("The isolated Teams test account is not available.");
  }
  const marker = `READY-${randomUUID().slice(0, 8).toUpperCase()}`;
  const message = `Reply with exactly this text and nothing else: ${marker}`;
  let stage = "start Chrome";
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  const page = await browser.newPage({ locale: "en-US" });
  try {
    stage = "Teams sign-in";
    await page.goto("https://teams.cloud.microsoft/v2/", { waitUntil: "domcontentloaded" });
    const loginId = page.locator('input[name="loginfmt"]');
    await loginId.waitFor({ timeout: 30000 });
    if (new URL(page.url()).hostname !== "login.microsoftonline.com") {
      throw new Error("Unexpected username form host.");
    }
    await loginId.fill(user);
    await page.getByRole("button", { name: "Next", exact: true }).click();
    const passwordInput = page.locator('input[name="passwd"]');
    await passwordInput.waitFor({ timeout: 30000 });
    if (new URL(page.url()).hostname !== "login.microsoftonline.com") {
      throw new Error("Unexpected password form host.");
    }
    await passwordInput.fill(password);
    await page.getByRole("button", { name: "Sign in", exact: true }).click();
    const keepSignedIn = page.getByRole("button", { name: "No", exact: true });
    const apps = page.getByRole("button", { name: "Apps", exact: true }).first();
    await keepSignedIn.or(apps).first().waitFor({ timeout: 45000 });
    if (new URL(page.url()).hostname === "login.microsoftonline.com" && await keepSignedIn.isVisible()) {
      await keepSignedIn.click();
    }
    stage = "Teams Apps navigation";
    await apps.waitFor({ timeout: 90000 });
    if (!teamsHosts.has(new URL(page.url()).hostname)) {
      throw new Error("Unexpected Teams app catalog host.");
    }
    await apps.click();
    stage = "Manage your apps";
    await page.getByRole("button", { name: "Manage your apps", exact: true }).click({ timeout: 30000 });
    stage = "locate the already installed test app";
    const app = page.getByText(appName, { exact: true }).first();
    await app.click({ timeout: 45000 });
    stage = "open the app's personal scope";
    await page.getByRole("button", { name: "Open", exact: true }).first().click({ timeout: 30000 });
    stage = "verify the chat recipient";
    const header = page.getByRole("heading", { name: appName, exact: true }).first();
    await header.waitFor({ timeout: 45000 });
    if (!teamsHosts.has(new URL(page.url()).hostname)) {
      throw new Error("Unexpected Teams conversation host.");
    }
    const composer = page.getByRole("textbox", { name: "Type a message", exact: true });
    const response = page.getByText(marker, { exact: true }).last();
    if (await response.isVisible()) {
      throw new Error("The expected reply was already present before sending.");
    }
    stage = "send one synthetic message";
    await composer.pressSequentially(message);
    await composer.press("Enter");
    stage = "receive the bot's reply";
    await response.waitFor({ timeout: 90000 });
    console.log("Teams verification passed: the intended installed bot returned the unique marker.");
  } catch (error) {
    await reportBrowserState(page);
    throw new Error(`Teams verification stopped at ${stage} (${error.name}).`);
  } finally {
    await browser.close();
  }
}

run().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
