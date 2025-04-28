// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
declare global {
  namespace NodeJS {
    interface ProcessEnv {
      APP_NAME_SUFFIX?: string;
      RESOURCE_SUFFIX?: string;
      TEAMSFX_ENV?: string;

      AZURE_SUBSCRIPTION_ID?: string;
      AZURE_RESOURCE_GROUP_NAME?: string;
      TEAMS_APP_UPDATE_TIME?: string;
      TEAMSFX_NEW_PROJECT_TYPE_NAME?: string;
      TEAMSFX_NEW_PROJECT_TYPE_EXTENSION?: string;
      APP_STUDIO_ENV?: string;

      SIDELOADING_SERVICE_SCOPE?: string;
      SIDELOADING_SERVICE_ENDPOINT?: string;

      KIOTA_BINARY_PATH?: string;
      AAD_APP_ACCESS_AS_USER_PERMISSION_ID?: string;

      AAD_APP_CLIENT_ID?: string;
      AAD_APP_OBJECT_ID?: string;
      AAD_APP_TENANT_ID?: string;
      AAD_APP_OAUTH_AUTHORITY_HOST?: string;
      AAD_APP_OAUTH_AUTHORITY?: string;
    }
  }
}

export {};
