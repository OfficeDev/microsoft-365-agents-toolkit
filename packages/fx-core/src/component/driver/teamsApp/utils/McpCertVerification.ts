// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { execFile } from "child_process";
import { promisify } from "util";
import { LocalMcpPrefix } from "../../../constants";
import { ODRProvider } from "../../../utils/odrProvider";

export const mcpCertDeps = {
  execFile: promisify(execFile),
};

/**
 * Verifies that all Local MCP plugin certificates in the plugin file are valid.
 * Returns true if all certs are valid or no Local MCP plugins exist, false otherwise.
 */
export async function verifyLocalMCPPluginCerts(pluginFile: AdmZip.IZipEntry): Promise<boolean> {
  const pluginContent = pluginFile.getData().toString();
  const pluginManifest = JSON.parse(pluginContent);
  if (!pluginManifest.runtimes || !Array.isArray(pluginManifest.runtimes)) {
    return true;
  }

  let allValidCerts = true;

  const localPluginRuntimes = pluginManifest.runtimes.filter(
    (runtime: { type: string }) => runtime.type === "LocalPlugin"
  );

  const localMcpPluginRuntimes = localPluginRuntimes.filter((runtime: any) =>
    runtime?.spec?.local_endpoint?.startsWith(LocalMcpPrefix)
  );

  if (localMcpPluginRuntimes.length === 0) {
    return true;
  }

  const servers = await ODRProvider.listServers();

  for (const runtime of localMcpPluginRuntimes) {
    const localEndpoint = (runtime as { spec?: { local_endpoint?: string } }).spec?.local_endpoint;

    if (!localEndpoint) continue;

    const mcpIdentifier = localEndpoint.substring(LocalMcpPrefix.length);
    const serverInfo = servers.find((x) => x.identifier === mcpIdentifier);

    if (!serverInfo) {
      continue;
    }

    const valid = await verifyPackageFamilyCertIsValid(serverInfo.packageFamily);

    if (!valid) {
      allValidCerts = false;
      break;
    }
  }

  return allValidCerts;
}

/**
 * Verifies that a package family certificate is valid (not a developer/self-signed cert).
 * Valid certs include: Store, System, Enterprise.
 * Invalid certs include: Developer (self-signed).
 */
export async function verifyPackageFamilyCertIsValid(packageName: string): Promise<boolean> {
  // Registry metadata must remain data, never part of the PowerShell script.
  const command =
    "Get-AppxPackage -ErrorAction Stop | " +
    "Where-Object { $_.PackageFamilyName -eq $env:ATK_MCP_PACKAGE_FAMILY } | " +
    "Select-Object -ExpandProperty SignatureKind";

  try {
    const { stdout } = await mcpCertDeps.execFile(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-Command", command],
      {
        shell: false,
        encoding: "utf8",
        env: { ...process.env, ATK_MCP_PACKAGE_FAMILY: packageName },
      }
    );
    return ["Store", "System", "Enterprise"].includes(stdout.trim());
  } catch (error) {
    console.error("Unable to get cert info for package name", error);
    return false;
  }
}
