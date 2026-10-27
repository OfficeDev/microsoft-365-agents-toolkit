const http = require("node:http");

function createDcrUpdateMock(initialEndpoint, metadata) {
  const initial = new URL(initialEndpoint);
  if (
    initial.protocol !== "https:" ||
    initial.username ||
    initial.password ||
    !metadata ||
    typeof metadata.resource !== "string" ||
    !Array.isArray(metadata.authorization_servers) ||
    !metadata.authorization_servers.length
  ) {
    throw new Error(
      "DCR mock requires a public HTTPS endpoint and real protected-resource metadata",
    );
  }
  const discovery = new URL("/.well-known/oauth-protected-resource", initial)
    .href;
  return http.createServer((request, response) => {
    const route = new URL(request.url, "http://localhost").pathname;
    response.setHeader("Content-Type", "application/json");
    response.setHeader("Cache-Control", "no-store");
    if (request.method === "GET" && route === "/health") {
      response.end(
        JSON.stringify({ status: "ready", fixture: "dcr-update-mock" }),
      );
    } else if (
      request.method === "GET" &&
      [
        "/.well-known/oauth-protected-resource",
        "/.well-known/oauth-protected-resource" + initial.pathname,
      ].includes(route)
    ) {
      response.end(JSON.stringify(metadata));
    } else if (route === initial.pathname) {
      response.writeHead(401, {
        "WWW-Authenticate": `Bearer resource_metadata="${discovery}"`,
      });
      response.end(JSON.stringify({ error: "unauthorized" }));
    } else {
      response.writeHead(404);
      response.end(JSON.stringify({ error: "not_found" }));
    }
  });
}

module.exports = { createDcrUpdateMock };

if (require.main === module) {
  (async () => {
    const [initial, endpoint, portText] = process.argv.slice(2);
    const port = Number(portText);
    if (!Number.isInteger(port) || port < 1 || port > 65535)
      throw new Error("Invalid mock listener port");
    const real = new URL(endpoint);
    if (real.protocol !== "https:")
      throw new Error("Real MCP endpoint must use HTTPS");
    const response = await fetch(
      new URL("/.well-known/oauth-protected-resource", real),
      { signal: AbortSignal.timeout(30000) },
    );
    if (!response.ok)
      throw new Error("Real MCP metadata read failed: HTTP " + response.status);
    const server = createDcrUpdateMock(initial, await response.json());
    server.on("error", (error) => {
      console.error("DCR mock listener failed:", error.code);
      process.exitCode = 1;
    });
    server.listen(port, "127.0.0.1", () =>
      console.log("DCR update mock ready on port", port),
    );
  })().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
