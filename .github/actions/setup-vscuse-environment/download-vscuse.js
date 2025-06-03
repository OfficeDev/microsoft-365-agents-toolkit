import fs from "fs";
import path from "path";
import fetch from "node-fetch";
import { createAppAuth } from "@octokit/auth-app";
import { Octokit } from "@octokit/core";

const appId = process.env.GH_APP_ID;
const privateKey = process.env.GH_APP_PRIVATE_KEY.replace(/\\n/g, "\n");
const owner = "microsoft"; // Replace with actual repo owner
const repo = "vscuse-doc"; // Replace with actual repo name

// Get target tag and store path from command line arguments
const targetTag = process.argv[2] || "latest";

// Handle the store path with proper quotation
let storePath = process.argv[3] || ".";
// Remove any quotes that might have been passed
storePath = storePath.replace(/^["']|["']$/g, '');

// Fix Windows paths
if (process.platform === 'win32') {
    // Ensure proper Windows path format
    storePath = path.resolve(storePath);
}

(async () => {
  try {
    console.log("Input store path:", process.argv[3]);
    console.log("Normalized store path:", storePath);
    console.log("🔐 Authenticating GitHub App...");
    const auth = createAppAuth({ appId, privateKey });
    const appAuth = await auth({ type: "app" });

    console.log("📥 Fetching installations...");
    const installationsRes = await fetch("https://api.github.com/app/installations", {
      headers: {
        Authorization: `Bearer ${appAuth.token}`,
        Accept: "application/vnd.github+json",
      },
    });
    const installations = await installationsRes.json();
    const installation = installations.find(i => i.account.login.toLowerCase() === owner.toLowerCase());
    if (!installation) throw new Error(`No installation found for account: ${owner}`);

    console.log(`🔑 Installation ID: ${installation.id}`);
    const installationAuth = await auth({
      type: "installation",
      installationId: installation.id,
    });

    const octokit = new Octokit({ auth: installationAuth.token });

    console.log("📄 Fetching releases...");
    const releases = await octokit.request("GET /repos/{owner}/{repo}/releases", {
      owner,
      repo,
    });

    console.log(`🎯 Found ${releases.data.length} releases`);
    for (const r of releases.data) {
      console.log(`- tag: ${r.tag_name}, draft: ${r.draft}, prerelease: ${r.prerelease}`);
    }

    let release;
    if (targetTag === "latest") {
      release = releases.data[0];
      if (!release) throw new Error("No releases found");
      console.log(`✅ Using latest release: ${release.tag_name}`);
    } else {
      release = releases.data.find(r => r.tag_name === targetTag);
      if (!release) throw new Error(`Release with tag '${targetTag}' not found`);
      console.log(`✅ Using specified release: ${release.tag_name}`);
    }

    release.assets.forEach(asset => {
      console.log(`- asset: ${asset.name} (${asset.browser_download_url})`);
    });

    const exeAsset = release.assets.find(asset => asset.name.endsWith(".exe"));
    if (!exeAsset) throw new Error("No .exe file found in release assets");

    console.log(`⬇️ Downloading ${exeAsset.name}...`);
    const res = await octokit.request('GET /repos/{owner}/{repo}/releases/assets/{asset_id}', {
      owner,
      repo,
      asset_id: exeAsset.id,
      headers: {
        'Accept': 'application/octet-stream'
      }
    });

    if (!res.data) throw new Error(`Failed to download asset`);

    console.log('storage path:', storePath);
    // Ensure the store path exists
    const normalizedPath = path.normalize(storePath);
    console.log('normalized path:', normalizedPath);
    await fs.promises.mkdir(normalizedPath, { recursive: true });
    const dest = path.join(normalizedPath, exeAsset.name);
    console.log('dest path:', dest);
    await fs.promises.writeFile(dest, Buffer.from(res.data));

    console.log(`✅ Download complete: ${dest}`);
  } catch (err) {
    console.error("❌ Error:", err.message);
    console.error(err.stack);
    process.exit(1);
  }
})();
