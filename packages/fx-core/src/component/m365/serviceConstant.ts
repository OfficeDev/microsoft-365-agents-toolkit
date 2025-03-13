// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const MosServiceEndpoint = "{{SERVICE_ENDPOINT_PLACEHOLDER}}";
export const MosServiceScope = "{{SERVICE_SCOPE_PLACEHOLDER}}";

export interface MOS3Api {
  key: string;
  method: "GET" | "POST" | "PUT" | "DELETE";
  path: RegExp;
}

export const MOS3ApiDefinitions: any = {
  GetConfigEnv: {
    key: "get_config_env",
    method: "GET",
    path: /^\/config\/v1\/environment$/,
  },
  PostPackageAddin: {
    key: "post_package_addin",
    method: "POST",
    path: /^\/dev\/v1\/users\/packages\/addins$/,
  },
  PostDevPackage: {
    key: "post_dev_package",
    method: "POST",
    path: /^\/dev\/v1\/users\/packages$/,
  },
  GetDevStatus: {
    key: "get_dev_status",
    method: "GET",
    path: /^\/dev\/v1\/users\/packages\/status\/[^\/]+/,
  },
  PostBuilderPackage: {
    key: "post_builder_package",
    method: "POST",
    path: /^\/builder\/v1\/users\/packages\?/,
  },
  GetBuilderStatus: {
    key: "get_builder_status",
    method: "GET",
    path: /^\/builder\/v1\/users\/packages\/status\/[^\/]+/,
  },
  PostDevPackageAcquisitions: {
    key: "post_dev_package_acquire",
    method: "POST",
    path: /^\/dev\/v1\/users\/packages\/acquisitions/,
  },
  GetShareInfo: {
    key: "get_share_info",
    method: "GET",
    path: /^\/marketplace\/v1\/users\/titles\/[^\/]+\/sharingInfo$/,
  },
  GetCatalogLaunchInfo: {
    key: "get_launch_info",
    method: "GET",
    path: /^\/catalog\/v1\/users\/titles\/launchInfo/,
  },
  DeleteCatalogAcquisitions: {
    key: "delete_catalog_acquisitions",
    method: "DELETE",
    path: /^\/catalog\/v1\/users\/acquisitions\/[^\/]+/,
  },
  GetLaunchInfoByTitle: {
    key: "get_launch_info_by_title",
    method: "GET",
    path: /^\/catalog\/v1\/users\/titles\/[^\/]+\/launchInfo/,
  },
  GetCatalogUITypes: {
    key: "get_catalog_ui_types",
    method: "GET",
    path: /^\/catalog\/v1\/users\/uitypes$/,
  },
};
