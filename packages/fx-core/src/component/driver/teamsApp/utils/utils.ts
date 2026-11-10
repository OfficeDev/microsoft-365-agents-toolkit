// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { includes } from "lodash";
import Mustache from "mustache";
import { getAppCapabilities, DevPortalAppResponse } from "../../../utils/appInfo";
import { getEnvironmentVariables } from "../../../utils/common";
import { WrapDriverContext } from "../../util/wrapUtil";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { TelemetryPropertyKey } from "./telemetry";
import { resolveManifestWithContext, ManifestType } from "../../../utils/envFunctionUtils";
import { DriverContext } from "../../interface/commonArgs";
export { RetryHandler } from "../../../../common/retryHandler";

export function getCustomizedKeys(prefix: string, manifest: any): string[] {
  let keys: string[] = [];
  for (const key in manifest) {
    if (manifest.hasOwnProperty(key)) {
      const value = manifest[key];
      if (typeof value === "object") {
        if (Array.isArray(value)) {
          value.map((item, index) => {
            keys = keys.concat(getCustomizedKeys(`${prefix}.${key}[${index}]`, item));
          });
        } else {
          keys = keys.concat(getCustomizedKeys(`${prefix}.${key}`, value));
        }
      } else if (typeof value === "string" && value.startsWith("{{config.manifest")) {
        keys.push(`${prefix}.${key}`);
      }
    }
  }
  return keys;
}

export function renderTemplate(manifestString: string, view: any): string {
  // Unesacped HTML
  Mustache.escape = (value) => value;
  manifestString = Mustache.render(manifestString, view);
  return manifestString;
}

export function isPersonalApp(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  const restrictedEntityIds: Array<string> = [
    "conversations",
    "recent",
    "about",
    "alltabs",
    "chat",
  ];
  if (!!appDefinition.staticTabs && appDefinition.staticTabs.length > 0) {
    return (
      appDefinition.staticTabs.filter((tab) => !restrictedEntityIds.includes(tab.entityId ?? ""))
        .length > 0
    );
  }

  return false;
}

export function isGroupApp(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  return (
    !!appDefinition.configurableTabs &&
    appDefinition.configurableTabs.length > 0 &&
    groupAppConfigured(appDefinition.configurableTabs[0])
  );
}

export function needTabAndBotCode(appDefinition: DevPortalAppResponse): boolean {
  return needTabCode(appDefinition) && needBotCode(appDefinition);
}

export function needTabCode(appDefinition: DevPortalAppResponse): boolean {
  return (isPersonalApp(appDefinition) || isGroupApp(appDefinition)) && !needBotCode(appDefinition);
}

export function isBot(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  return !!appDefinition.bots && appDefinition.bots.length > 0;
}

export function containsBotBasedMessageExtension(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  return (
    !!appDefinition.messagingExtensions &&
    appDefinition.messagingExtensions.length > 0 &&
    !!appDefinition.messagingExtensions[0].botId
  );
}
export function isBotBasedMessageExtension(appDefinition: DevPortalAppResponse): boolean {
  return containsBotBasedMessageExtension(appDefinition) && !isBot(appDefinition);
}

export function isBotAndBotBasedMessageExtension(appDefinition: DevPortalAppResponse): boolean {
  return isBot(appDefinition) && isBotBasedMessageExtension(appDefinition);
}

export function needBotCode(appDefinition: DevPortalAppResponse): boolean {
  return isBot(appDefinition) || isBotBasedMessageExtension(appDefinition);
}

function isApiBasedMessageExtension(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  const extension = appDefinition.messagingExtensions?.[0];
  if (!extension) return false;
  const serviceType =
    "composeExtensionType" in extension
      ? extension.composeExtensionType
      : "messagingExtensionServiceType" in extension
        ? extension.messagingExtensionServiceType
        : undefined;
  return typeof serviceType === "string" && serviceType.toLowerCase() === "apibased";
}

export function containsUnsupportedFeature(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  const hasScene = appDefinition?.meetingExtensionDefinition?.scenes?.length;
  const hasConnector = !!appDefinition?.connectors?.length;
  const hasActivies = appDefinition?.activities?.activityTypes?.length;

  return (
    !!hasScene ||
    !!hasConnector ||
    !!hasActivies ||
    hasMeetingExtension(app) ||
    isApiBasedMessageExtension(app)
  );
}

export function getFeaturesFromAppDefinition(appDefinition: DevPortalAppResponse): string[] {
  const features = [];
  const personalTab = "personal-tab";
  const groupTab = "group-tab";
  const bot = "bot";
  const messageExtension = "messaging-extension";

  if (isPersonalApp(appDefinition)) {
    features.push(personalTab);
  }

  if (isGroupApp(appDefinition)) {
    features.push(groupTab);
  }

  if (isBot(appDefinition)) {
    features.push(bot);
  }

  if (containsBotBasedMessageExtension(appDefinition)) {
    features.push(messageExtension);
  }

  return features;
}

export function hasMeetingExtension(app: DevPortalAppResponse): boolean {
  const appDefinition = getAppCapabilities(app);
  return (
    !!appDefinition.configurableTabs &&
    appDefinition.configurableTabs.length > 0 &&
    meetingExtensionConfigured(appDefinition.configurableTabs[0])
  );
}

type TabContext = { context?: string[] | null; scopes?: string[] | null };

const groupAppConfigured = (tab: TabContext) => {
  const validGroupAppContext =
    includes(tab.context, MeetingsContext.ChannelTab) ||
    includes(tab.context, MeetingsContext.PrivateChatTab);

  const validGroupAppScope =
    (!!tab.scopes && includeGroupChatScope(tab.scopes)) || includeTeamScope(tab.scopes);

  return validGroupAppScope && validGroupAppContext;
};

const meetingExtensionConfigured = (tab: TabContext) => {
  const validMeetingContext =
    includes(tab.context, MeetingsContext.SidePanel) ||
    includes(tab.context, MeetingsContext.DetailsTab) ||
    includes(tab.context, MeetingsContext.ChatTab);

  const validMeetingScope = !!tab.scopes && includeGroupChatScope(tab.scopes);

  return validMeetingScope && validMeetingContext;
};

const includeTeamScope = (scopes?: string[] | null): boolean => {
  return !!scopes?.find((scope) => scope.toLowerCase() === CommandScope.Team.toLowerCase());
};

const includeGroupChatScope = (scopes?: string[] | null): boolean => {
  return !!scopes?.find((scope) => scope.toLowerCase() === CommandScope.GroupChat.toLowerCase());
};

export enum CommandScope {
  Team = "team",
  Personal = "personal",
  GroupChat = "groupChat",
}

export enum MeetingsContext {
  ChannelTab = "channelTab",
  PrivateChatTab = "privateChatTab",
  SidePanel = "meetingSidePanel",
  ShareToStage = "meetingStage",
  DetailsTab = "meetingDetailsTab",
  ChatTab = "meetingChatTab",
}

export function normalizePath(path: string, useForwardSlash: boolean): string {
  return useForwardSlash ? path.replace(/\\/g, "/") : path;
}

export async function getResolvedManifest(
  content: string,
  path: string,
  manifestType: ManifestType,
  ctx: DriverContext
): Promise<Result<string, FxError>> {
  const vars = getEnvironmentVariables(content);
  let telemetryKey;
  switch (manifestType) {
    case ManifestType.ApiSpec:
      telemetryKey = TelemetryPropertyKey.customizedOpenAPIKeys;
      break;
    case ManifestType.PluginManifest:
      telemetryKey = TelemetryPropertyKey.customizedAIPluginKeys;
      break;
    case ManifestType.DeclarativeCopilotManifest:
      telemetryKey = TelemetryPropertyKey.customizedCopilotGptKeys;
      break;
    default:
      telemetryKey = TelemetryPropertyKey.customizedKeys;
      break;
  }

  if (ctx instanceof WrapDriverContext) {
    ctx.addTelemetryProperties({
      [telemetryKey]: vars.join(";"),
    });
  }

  return resolveManifestWithContext(content, ctx, manifestType, path);
}
