import fs from "fs";
import path from "path";
import { createHash } from "node:crypto";
import fetch from "node-fetch";
import { createAppAuth } from "@octokit/auth-app";
import { Octokit } from "@octokit/core";

const appId = process.env.GH_APP_ID;
const privateKey = process.env.GH_APP_PRIVATE_KEY.replace(/\\n/g, "\n");
const aadValidation = process.env.GITHUB_REF === "refs/heads/qfai/vscuse-aad-prerelease-validation";
// Do not publish validation wheels to vscuse-doc: its daily consumer takes releases[0].
const owner = aadValidation ? "1openwindow" : "microsoft";
const repo = aadValidation ? "vsc-use" : "vscuse-doc";

// Get target tag and store path from command line arguments
const targetTag = process.argv[2] || "latest";
const expectedSha256 = process.env.VSCUSE_SHA256 || "";
const candidateCommit = "e8f2e0ecbf52353584358907b8f6b83ec00c4359";

// Handle the store path with proper quotation
let storePath = process.argv[3] || ".";
// Remove any quotes that might have been passed
storePath = storePath.replace(/^["']|["']$/g, '');

// Fix Windows paths
if (process.platform === 'win32') {
    // Ensure proper Windows path format
    storePath = path.resolve(storePath);
}

console.log("GH_APP_PRIVATE_KEY length:", process.env.GH_APP_PRIVATE_KEY ? process.env.GH_APP_PRIVATE_KEY.length : 0);

(async () => {
  try {
    if (aadValidation && (targetTag !== "v1.0.3rc1" ||
        !/^[a-f0-9]{64}$/.test(expectedSha256) || expectedSha256.length !== 64)) {
      throw new Error("AAD validation requires v1.0.3rc1 and the expected wheel SHA256");
    }
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

    let release;
    if (aadValidation) {
      const result = await octokit.request("GET /repos/{owner}/{repo}/releases/tags/{tag}", {
        owner,
        repo,
        tag: targetTag,
      });
      release = result.data;
      if (release.tag_name !== targetTag || !release.prerelease || release.draft ||
          release.target_commitish !== candidateCommit) {
        throw new Error("AAD validation requires the exact published prerelease and source commit");
      }
    } else {
      console.log("📄 Fetching releases...");
      const releases = await octokit.request("GET /repos/{owner}/{repo}/releases", {
        owner,
        repo,
      });
      console.log(`🎯 Found ${releases.data.length} releases`);
      for (const r of releases.data) {
        console.log(`- tag: ${r.tag_name}, draft: ${r.draft}, prerelease: ${r.prerelease}`);
      }
      if (targetTag === "latest") {
        release = releases.data[0];
        if (!release) throw new Error("No releases found");
        console.log(`✅ Using latest release: ${release.tag_name}`);
      } else {
        release = releases.data.find(r => r.tag_name === targetTag);
        if (!release) throw new Error(`Release with tag '${targetTag}' not found`);
        console.log(`✅ Using specified release: ${release.tag_name}`);
      }
    }

    release.assets.forEach(asset => {
      console.log(`- asset: ${asset.name} (${asset.browser_download_url})`);
    });

    const exeAsset = release.assets.find(asset => asset.name.endsWith(".whl"));
    if (!exeAsset) throw new Error("No .exe file found in release assets");
    if (aadValidation) {
      const wheels = release.assets.filter(asset => asset.name.endsWith(".whl"));
      if (wheels.length !== 1 || path.basename(exeAsset.name) !== exeAsset.name ||
          !exeAsset.name.startsWith(`vscuse-${targetTag.slice(1)}-`) ||
          exeAsset.digest !== `sha256:${expectedSha256}`) {
        throw new Error("Expected one version-matched vscuse wheel with a SHA256 digest");
      }
    }

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
    if (aadValidation &&
        createHash("sha256").update(Buffer.from(res.data)).digest("hex") !== expectedSha256) {
      throw new Error("Prerelease wheel SHA256 mismatch");
    }

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

