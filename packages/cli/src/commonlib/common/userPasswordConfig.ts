// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import dotenv from "dotenv";

dotenv.config();

export const client_id = "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0";

export const AZURE_ACCOUNT_NAME = process.env.AZURE_ACCOUNT_NAME; // Login with the azure username
export const AZURE_ACCOUNT_OBJECT_ID = process.env.AZURE_ACCOUNT_OBJECT_ID;
export const AZURE_ACCOUNT_PASSWORD = process.env.AZURE_ACCOUNT_PASSWORD;
export const AZURE_SERVICE_PRINCIPAL_ID = process.env.AZURE_SERVICE_PRINCIPAL_ID; // Login with the service principal
export const AZURE_SERVICE_PRINCIPAL_SECRET = process.env.AZURE_SERVICE_PRINCIPAL_SECRET;
export const AZURE_SUBSCRIPTION_ID = process.env.AZURE_SUBSCRIPTION_ID;
export const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
export const M365_ACCOUNT_NAME = process.env.M365_ACCOUNT_NAME;
export const M365_ACCOUNT_PASSWORD = process.env.M365_ACCOUNT_PASSWORD;
export const M365_TENANT_ID = process.env.M365_TENANT_ID;
