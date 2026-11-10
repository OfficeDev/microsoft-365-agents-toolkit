// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError, UserError } from "@microsoft/teamsfx-api";
import axios from "axios";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

const retryableStatuses = new Set([429, 502, 503, 504]);
const retryableTransportCodes = new Set(["ECONNRESET", "ETIMEDOUT", "ECONNABORTED", "EAI_AGAIN"]);
const retryDeadlineMs = 5 * 60 * 1000;

/**
 * Use "never" for directory creation and final package submission.
 * Retried HTTP calls must also set a finite request timeout.
 */
export async function requestSkillOperation<T>(
  request: () => Promise<T>,
  replaySafe: boolean | "never" = true
): Promise<Result<T, FxError>> {
  const deadline = Date.now() + retryDeadlineMs;
  for (let attempt = 0; ; attempt++) {
    try {
      return ok(await request());
    } catch (error: unknown) {
      const status = axios.isAxiosError(error) ? error.response?.status : undefined;
      const transportCode = axios.isAxiosError(error) ? error.code : undefined;
      const retryable =
        replaySafe !== "never" &&
        (status === 429 ||
          (replaySafe === true &&
            ((status !== undefined && retryableStatuses.has(status)) ||
              (status === undefined &&
                transportCode !== undefined &&
                retryableTransportCodes.has(transportCode)))));
      const fallback = 1000 * 2 ** attempt + Math.floor(Math.random() * 250);
      const header = axios.isAxiosError(error)
        ? error.response?.headers?.["retry-after"]
        : undefined;
      const delay = retryAfterMs(header) ?? fallback;
      if (retryable && attempt < 5 && Date.now() + delay < deadline) {
        await new Promise<void>((resolve) => setTimeout(resolve, delay));
        continue;
      }

      const authorization = status === 401 || status === 403;
      const conflict = status === 409 || status === 412;
      const name = authorization
        ? "SkillUploadAuthorizationFailed"
        : conflict
          ? "SkillUploadConflict"
          : retryable
            ? "SkillUploadRetryExhausted"
            : "SkillUploadRequestFailed";
      const key = authorization
        ? "error.m365.skills.authorization"
        : conflict
          ? "error.m365.skills.conflict"
          : retryable
            ? "error.m365.skills.retryExhausted"
            : "error.m365.skills.requestFailed";
      const options = {
        source: "M365",
        name,
        message: getDefaultString(key),
        displayMessage: getLocalizedString(key),
        // Axios errors carry bearer headers and credential-bearing URLs.
        telemetryProperties: {
          status: status?.toString() ?? "transport",
          attempts: (attempt + 1).toString(),
          transportCode:
            transportCode && retryableTransportCodes.has(transportCode) ? transportCode : "",
        },
      };
      return err(
        authorization ||
          conflict ||
          (status !== undefined && status >= 400 && status < 500 && status !== 429)
          ? new UserError(options)
          : new SystemError(options)
      );
    }
  }
}

function retryAfterMs(value: unknown): number | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const text = String(value).trim();
  if (/^\d+$/.test(text)) {
    const seconds = Number(text);
    return Number.isFinite(seconds) ? seconds * 1000 : undefined;
  }
  if (!/^[A-Za-z]{3},/.test(text)) return undefined;
  const date = Date.parse(text);
  return Number.isFinite(date) ? Math.max(0, date - Date.now()) : undefined;
}
