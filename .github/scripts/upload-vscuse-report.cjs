const fs = require("node:fs");
const path = require("node:path");

const apimUploadEndpoint =
  "https://apim-ai-atktest1022.azure-api.net/vscuse-report-upload";
const maxFileBytes = 100 * 1024 * 1024;
const contentTypes = {
  ".css": "text/css",
  ".gif": "image/gif",
  ".htm": "text/html",
  ".html": "text/html",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "text/javascript",
  ".json": "application/json",
  ".mp4": "video/mp4",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain",
  ".webm": "video/webm",
  ".webp": "image/webp",
  ".zip": "application/zip",
};

async function uploadReport(
  reportPath,
  guid,
  { key, endpoint = apimUploadEndpoint },
) {
  if (!key?.trim()) {
    throw new Error("VSCUSE_REPORT_APIM_KEY is not set.");
  }
  if (!/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(guid)) {
    throw new Error("Expected a report GUID.");
  }

  const source = fs.lstatSync(reportPath);
  if (!source.isFile() && !source.isDirectory()) {
    throw new Error("Expected a report file or directory.");
  }
  const files = source.isFile()
    ? [{ localPath: reportPath, blobPath: "index.html", size: source.size }]
    : fs
        .readdirSync(reportPath, { recursive: true, withFileTypes: true })
        .filter((entry) => !entry.isDirectory())
        .map((entry) => {
          const localPath = path.join(entry.parentPath, entry.name);
          const blobPath = path
            .relative(reportPath, localPath)
            .split(path.sep)
            .join("/");
          if (!entry.isFile()) {
            throw new Error(`Not a regular report file: ${blobPath}`);
          }
          return { localPath, blobPath, size: fs.statSync(localPath).size };
        });
  if (
    !files.some(({ blobPath, size }) => blobPath === "index.html" && size > 0)
  ) {
    throw new Error("Report has no nonempty index.html.");
  }
  for (const { blobPath, size } of files) {
    if (size > maxFileBytes) {
      throw new Error(`Report file exceeds 100 MiB: ${blobPath}`);
    }
  }

  for (const { localPath, blobPath, size } of files) {
    const encodedPath = blobPath.split("/").map(encodeURIComponent).join("/");
    const body = size === 0 ? undefined : fs.createReadStream(localPath);
    try {
      const response = await fetch(`${endpoint}/${guid}/${encodedPath}`, {
        method: "PUT",
        headers: {
          "Ocp-Apim-Subscription-Key": key,
          "Content-Length": String(size),
          "Content-Type":
            contentTypes[path.extname(blobPath).toLowerCase()] ??
            "application/octet-stream",
        },
        body,
        duplex: "half",
        redirect: "manual",
        signal: AbortSignal.timeout(120_000),
      });
      await response.body?.cancel();
      if (response.status !== 201) {
        throw new Error(
          `APIM report upload failed (HTTP ${response.status}): ${blobPath}`,
        );
      }
    } finally {
      body?.destroy();
    }
  }
}

if (require.main === module) {
  uploadReport(process.argv[2], process.argv[3], {
    key: process.env.VSCUSE_REPORT_APIM_KEY,
  }).catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}

module.exports = { uploadReport };
