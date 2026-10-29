// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { featureFlagManager, FeatureFlags } from "../common/featureFlags";
import { DevPortalClient, devPortalClient } from "./devPortalClient";
import { TeamsDevPortalClient, legacyTeamsDevPortalClient } from "./teamsDevPortalClient";
import { AppResponse } from "./devPortalModels";
import { AppDefinition } from "../component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { AppUser } from "../component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { APP_STUDIO_API_NAMES } from "../component/driver/teamsApp/constants";

const unsupportedNewApi = (): never => {
  throw new Error("The operation requires the new Developer Portal APIs.");
};
const unsupportedCreateApp: DevPortalClient["createApp"] = unsupportedNewApi;
const unsupportedUpdateApp: DevPortalClient["updateApp"] = unsupportedNewApi;
// The target describes the real public union contract. The proxy binds operations to the
// selected client; own test overrides are mirrored to both clients and restored below.
const facadeState: { regionEndpoint?: string } = {};
const facadeModel = Object.assign(facadeState, {
  setRegionEndpoint: devPortalClient.setRegionEndpoint,
  setRegionEndpointByToken: devPortalClient.setRegionEndpointByToken,
  getEndpoint: devPortalClient.getEndpoint,
  createRequesterWithToken: devPortalClient.createRequesterWithToken,
  wrapException: devPortalClient.wrapException,
  wrapResponse: devPortalClient.wrapResponse,
  handleBotFrameworkError: devPortalClient.handleBotFrameworkError,
  importApp: devPortalClient.importApp,
  createApp: unsupportedCreateApp,
  updateApp: unsupportedUpdateApp,
  getApp(token: string, appId: string) {
    return getActiveTeamsDevPortalClient().getApp(token, appId);
  },
  listApps(token: string) {
    return getActiveTeamsDevPortalClient().listApps(token);
  },
  getUserList(token: string, appId: string) {
    return getActiveTeamsDevPortalClient().getUserList(token, appId);
  },
  checkUser(app: AppDefinition | AppResponse, user: AppUser): boolean {
    return "appProfile" in app
      ? devPortalClient.checkUser(app, user)
      : legacyTeamsDevPortalClient.checkUser(app, user);
  },
  async getAppLastModified(token: string, resourceId: string): Promise<string> {
    const client = getActiveTeamsDevPortalClient();
    const updatedAt =
      client instanceof DevPortalClient
        ? (await client.listApps(token)).find((app) => app.appId === resourceId)?.updatedAt
        : (await client.getApp(token, resourceId)).updatedAt;
    if (
      typeof updatedAt !== "string" ||
      !updatedAt.trim() ||
      !Number.isFinite(Date.parse(updatedAt))
    ) {
      throw client.wrapException(
        new Error("The app response does not contain a valid last-modified timestamp."),
        APP_STUDIO_API_NAMES.GET_APP
      );
    }
    return updatedAt;
  },
  getBotId: devPortalClient.getBotId,
  getAppPackage(token: string, appId: string) {
    return getActiveTeamsDevPortalClient().getAppPackage(token, appId);
  },
  deleteApp: devPortalClient.deleteApp,
  checkExistsInTenant: devPortalClient.checkExistsInTenant,
  publishTeamsApp: devPortalClient.publishTeamsApp,
  publishTeamsAppUpdate: devPortalClient.publishTeamsAppUpdate,
  getStaggedApp: devPortalClient.getStaggedApp,
  checkPermission: devPortalClient.checkPermission,
  removePermission: devPortalClient.removePermission,
  grantPermission: devPortalClient.grantPermission,
  partnerCenterAppPackageValidation: devPortalClient.partnerCenterAppPackageValidation,
  submitAppValidationRequest: devPortalClient.submitAppValidationRequest,
  getAppValidationRequestList: devPortalClient.getAppValidationRequestList,
  getAppValidationById: devPortalClient.getAppValidationById,
  getSideloadingStatus: devPortalClient.getSideloadingStatus,
  createApiKeyRegistration: devPortalClient.createApiKeyRegistration,
  getApiKeyRegistrationById: devPortalClient.getApiKeyRegistrationById,
  updateApiKeyRegistration: devPortalClient.updateApiKeyRegistration,
  getOauthRegistrationById: devPortalClient.getOauthRegistrationById,
  createOauthRegistration: devPortalClient.createOauthRegistration,
  updateOauthRegistration: devPortalClient.updateOauthRegistration,
  getBotRegistration: devPortalClient.getBotRegistration,
  listBots: devPortalClient.listBots,
  deleteBot: devPortalClient.deleteBot,
  createBotRegistration: devPortalClient.createBotRegistration,
  updateMessageEndpoint: devPortalClient.updateMessageEndpoint,
  updateBotRegistration: devPortalClient.updateBotRegistration,
  createAADApp: devPortalClient.createAADApp,
});
const facadeTarget: typeof facadeModel = Object.create(facadeModel);
const concreteDescriptors = new Map<
  PropertyKey,
  [PropertyDescriptor | undefined, PropertyDescriptor | undefined]
>();

export function isUsingNewDeveloperPortalApis(): boolean {
  return featureFlagManager.getBooleanValue(FeatureFlags.NewDeveloperPortalApis);
}

export function getActiveTeamsDevPortalClient(): TeamsDevPortalClient | DevPortalClient {
  return isUsingNewDeveloperPortalApis() ? devPortalClient : legacyTeamsDevPortalClient;
}

export const teamsDevPortalClient = new Proxy(facadeTarget, {
  get(target, property, receiver) {
    if (concreteDescriptors.has(property)) {
      return Reflect.get(target, property, receiver);
    }
    if (property === "checkUser") return target.checkUser;
    const activeClient = getActiveTeamsDevPortalClient();
    const value = Reflect.get(activeClient, property, activeClient);
    if (value === undefined) {
      return Reflect.get(target, property, receiver);
    }
    return typeof value === "function" ? value.bind(activeClient) : value;
  },
  set(target, property, value, receiver) {
    if (
      Object.prototype.hasOwnProperty.call(target, property) ||
      typeof Reflect.get(getActiveTeamsDevPortalClient(), property) === "function"
    ) {
      return Reflect.set(target, property, value, receiver);
    }
    return Reflect.set(getActiveTeamsDevPortalClient(), property, value);
  },
  defineProperty(target, property, descriptor) {
    if (!concreteDescriptors.has(property)) {
      concreteDescriptors.set(property, [
        Object.getOwnPropertyDescriptor(legacyTeamsDevPortalClient, property),
        Object.getOwnPropertyDescriptor(devPortalClient, property),
      ]);
    }
    Object.defineProperty(legacyTeamsDevPortalClient, property, descriptor);
    Object.defineProperty(devPortalClient, property, descriptor);
    return Reflect.defineProperty(target, property, descriptor);
  },
  deleteProperty(target, property) {
    const descriptors = concreteDescriptors.get(property);
    if (descriptors) {
      const clients = [legacyTeamsDevPortalClient, devPortalClient];
      descriptors.forEach((descriptor, index) => {
        const client = clients[index];
        if (descriptor) {
          Object.defineProperty(client, property, descriptor);
        } else {
          Reflect.deleteProperty(client, property);
        }
      });
      concreteDescriptors.delete(property);
    }
    return Reflect.deleteProperty(target, property);
  },
});
