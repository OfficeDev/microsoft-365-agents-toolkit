// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { isMap, isScalar, isSeq, parseDocument } from "yaml";
import { getLocalizedString } from "../../../common/localizeUtils";
import authTemplates from "./assets/openApiAuth.json";

export interface AuthRegistration {
  authName: string;
  authType: "apiKey" | "oauth2";
  registrationIdEnvName: string;
}

function invalidYaml(): Result<never, FxError> {
  return err(
    new SystemError({
      source: "Scaffold",
      name: "OpenApiAuthYamlInvalid",
      message: getLocalizedString("core.v4.scaffold.openApiAuthYamlInvalid"),
    })
  );
}

export function injectOpenApiAuthActions(
  yaml: string,
  registrations: AuthRegistration[],
  apiSpecPath: string
): Result<string, FxError> {
  if (registrations.length === 0) return ok(yaml);
  const document = parseDocument(yaml, { intAsBigInt: true });
  if (document.errors.length > 0) return invalidYaml();
  const provision = document.get("provision", true);
  if (!isSeq(provision)) return invalidYaml();
  const anchor = provision.items.findIndex(
    (item) => isMap(item) && item.get("uses") === "teamsApp/zipAppPackage"
  );
  let index = anchor < 0 ? provision.items.length : anchor;
  for (const registration of registrations) {
    const actionDocument = parseDocument(authTemplates[registration.authType].join("\n"));
    const action = actionDocument.contents;
    if (!isMap(action)) return invalidYaml();
    const firstKey = action.items[0]?.key;
    if (isScalar(firstKey) && firstKey.commentBefore) {
      action.commentBefore = firstKey.commentBefore;
      firstKey.commentBefore = undefined;
    }
    action.setIn(["with", "name"], actionDocument.createNode(registration.authName));
    action.setIn(["with", "apiSpecPath"], actionDocument.createNode(apiSpecPath));
    action.setIn(
      [
        "writeToEnvironmentFile",
        registration.authType === "apiKey" ? "registrationId" : "configurationId",
      ],
      actionDocument.createNode(registration.registrationIdEnvName)
    );
    provision.items.splice(index++, 0, action);
  }
  provision.flow = false;
  return ok(document.toString({ lineWidth: 0 }));
}
