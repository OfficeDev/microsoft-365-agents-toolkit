# Contributing

Welcome and thank you for your interest in contributing to **TeamsFX SDK**! Before contributing to this project, please review this document for policies and procedures which will ease the contribution and review process for everyone. If you have questions, please raise your issue on github.

# Setting up development environment

Make sure you have installed following prerequisites.
To test the SDK, you'd better install [Visual Studio Code](https://code.visualstudio.com/) and [TeamsFx VS Code Extension](https://github.com/OfficeDev/TeamsFx/tree/main/packages/vscode-extension).

## Prerequisites

- Git
- Node 18 or higher (Node 18 is recommended)
- PNPM 8 or higher

# Building SDK

1. Clone this repo locally. (`git clone https://github.com/OfficeDev/TeamsFx.git`)
2. Open a terminal and move into your local copy. (`cd TeamsFx`)
3. Because the monorepo is managed by PNPM, you need to setup at the first time. (`npm run setup` or `pnpm install && npm run build` in root folder) All dependencies will be installed.

# Debugging SDK

Use [pnpm-pack](https://pnpm.io/cli/pack) to test the SDK.
If you use a project created by TeamsFx VS Code extension, remove the `@microsoft/teamsfx` in package.json first.

```bash
cd packages/sdk                             # go into the SDK package directory
pnpm pack                                    # create local packed file
cd ../test-project-using-sdk                # go into test project directory.
npm install ../microsoft-teamsfx-x.x.x.tgz  # install the local built package
```

Run `npm run build` and `npm pack` under `packages/sdk` after any updates of SDK.

## Add or remove dependency for SDK

run `pnpm install XXX` to add dependency, run `pnpm remove XXX` to remove dependency.

# Running test cases

All tests are under `test/` folder. The filename ends with ".spec.ts".

E2E tests need environment variables that are configured in GitHub Action. To run SDK E2E tests locally, environment variables are required to be set manually in the `.env` file under `packages/sdk`.

- Use `npm run test:unit` to run all unit tests.
- Use `npm run test:e2e` to run all e2e tests.

## Test folder structure

- test/unit/: unit tests for both browser and NodeJS.
- test/unit/browser/: unit tests for browser only.
- test/unit/node/: unit tests for NodeJS only.
- test/e2e/browser/: end-to-end tests related to Azure resources for browser only.
- test/e2e/node/: end-to-end tests related to Azure resources for NodeJS only.

## Prepare values for required environment variables to run E2E test cases

All the environment variables are extracted by the function `extractIntegrationEnvVariables()` defined in `helper.browser.ts` and `helper.ts` under `test/` folder.

### SDK_INTEGRATION_TEST_ACCOUNT

1. Open Teams Toolkit, and sign into M365 by clicking the `Sign in to M365` under the `ACCOUNTS` section from sidebar with your test account. After you signed in, create a new Teams `React with Fluent UI template` app. Start debugging the project by hitting the `F5` key in the Visual Studio Code.

2. Set the environment variable `SDK_INTEGRATION_TEST_ACCOUNT` with the test account name and password, separated by semicolons.

   ```
   # AAD Account (NodeJS & browser)
   SDK_INTEGRATION_TEST_ACCOUNT = {TEST_ACCOUNT_NAME};{TEST_ACCOUNT_PASSWORD}
   ```

### SDK_INTEGRATION_TEST_AAD

1. Open Azure Portal with your test account, go to `App Registrations` under `Azure Active Directory`, and find the app created and debugged in the previous section.

2. Click on the app name, and then go to `Authentication` under the `Manage` section from sidebar.

3. **Uncheck** Access tokens and **check** ID tokens under the `Implicit grant and hybrid flows` section. Set `Allow public client flows` to **Yes**.

4. Go to `API Permissions` under the `Manage` section from sidebar, and click `add a permission`. Select `Microsoft Graph` API and then select `Delegated permissions`. Scroll down to the bottom and expand `User`, and check `User.Read.All`. Permission of `User.Read.All` requires Admin consent.

5. Navigate to `Overview`, `AAD_TENANT_ID` and `AAD_CLIENT_ID` can be found under `Essentials` section.

6. `AAD_AUTHORITY_HOST` is `https://login.microsoftonline.com`.

7. `AAD_CLIENT_SECRET` can be found under `.fx/states/local.userdata`. Click `Decrypt secret` to get the value.

8. Follow the [link](https://docs.microsoft.com/en-us/azure/marketplace/find-tenant-object-id#find-user-object-id) to find `USER_OBJECT_ID` .

9. To setup the certificate, execute following commands to generate self signed certificates for test. The command requires OpenSSL.

   ```shell
   # Skip the Export Password with ENTER
   openssl req -x509 -newkey rsa:4096 -keyout PrivateKey.pem -out Cert.pem -sha256 -days 365 -nodes -subj "//CN=SdkIntegrationTest"

   openssl pkcs12 -export -out keyStore.pfx -inkey PrivateKey.pem -in Cert.pem

   openssl pkcs12 -in keyStore.pfx -out keyStore.pem -nodes
   ```

10. In the directory that executes step 9's commands, execute following js script. `AAD_CERTIFICATE_CONTENT` is the contents of `certificate.txt`.

    ```javascript
    const fs = require("fs");
    var cert_key = fs.readFileSync("keyStore.pem", "utf8");
    cert_key = cert_key.replace(/\r/g, "");
    fs.writeFileSync("certificate.txt", JSON.stringify(cert_key));
    ```

11. Go to `Certificates & secrets` of the app under `App registrations`, and upload `Cert.pem` by clicking `Upload certificate` under `Certificate` section.

12. Set the environment variable `SDK_INTEGRATION_TEST_AAD` with the following sequence, separated by semicolons.

    ```
    # AAD Application (NodeJS & browser)
    SDK_INTEGRATION_TEST_AAD = {AAD_AUTHORITY_HOST};{AAD_TENANT_ID};{USER_OBJECT_ID};{AAD_CLIENT_ID};{AAD_CLIENT_SECRET};{AAD_CERTIFICATE_CONTENT}
    ```

### SDK_INTEGRATION_TEST_API_CERTPROVIDER

1. Execute following commands to generate self signed certificates for test. The command requires OpenSSL.

   ```shell
   # Generate certs to for test server that supports certificate auth
   openssl req -x509 -newkey rsa:4096 -nodes -sha256 -days 365 -keyout server_key.pem -out server_cert.pem  -subj "//CN=localhost"

   # Generate client cert (both PEM and PFX) to test certificate auth support. Press "ENTER" when asked for "Export Password".
   openssl req -newkey rsa:4096 -sha256 -keyout client_key.pem -out client_csr.pem -nodes -days 365 -subj "//CN=test client"

   openssl x509 -req -in client_csr.pem -CA server_cert.pem -CAkey server_key.pem -out client_cert.pem -set_serial 01 -days 365

   openssl pkcs12 -inkey client_key.pem -in client_cert.pem -certfile server_cert.pem -export -out client_pfx.pfx

   # Create password protected private key file
   openssl rsa -aes256 -in client_key.pem -out client_key_encrypted.pem

   # Create password protected pfx file. Record the cert password for use in step 2.
   openssl pkcs12 -inkey client_key.pem -in client_cert.pem -certfile server_cert.pem -export -out client_pfx_encrypted.pfx
   ```

2. In the directory that executes step 1's commands, execute following js script to generate values to be configured to the environment variable. You need to fill in the cert password you provided in step 1 before running the script.

   ```javascript
   const fs = require("fs");

   let certs = {};
   certs.serverCert = fs.readFileSync("server_cert.pem", "utf8");
   certs.serverKey = fs.readFileSync("server_key.pem", "utf8");
   certs.clientCert = fs.readFileSync("client_cert.pem", "utf8");
   certs.clientKey = fs.readFileSync("client_key.pem", "utf8");
   certs.clientKeyEncrypted = fs.readFileSync("client_key_encrypted.pem", "utf8");
   certs.clientPfx = Buffer.from(fs.readFileSync("client_pfx.pfx")).toString("base64");
   certs.clientPfxEncrypted = Buffer.from(fs.readFileSync("client_pfx_encrypted.pfx")).toString(
     "base64"
   );
   certs.passphrase = ""; // fill your password before executing the script
   certs.clientCN = "test client"; // update the value if you specifies another CN when generating test certs

   fs.writeFileSync("certs.json", JSON.stringify(certs));
   ```

3. Copy the contents of `certs.json` to environment variable `SDK_INTEGRATION_TEST_API_CERTPROVIDER`.

# Style Guidelines

This project use eslint and prettier to check format and code style.
Use `npm run format` to fix format issues and `npm run lint` to check lint issues.

# Pull Request Process

1. Check out a new branch from "dev".
2. Update code in correct place. [Supporting Browser and NodeJS](#supporting-browser-and-nodejs)
3. Make sure modified codes are covered by unit tests. [Running test cases](#running-test-cases)
4. Ensure code style check has no error. [Style Guidelines](#style-guidelines)
5. Add comment for public class/method. Please check [comment template](API_COMMENT.md) for details.
6. Merge your changes to "dev" branch.
7. At least one approve from code owners is required.

## Supporting Browser and NodeJS

1. If a new class behaves differently under two environments. Create a new file named xxx.browser.ts that works only in browser and xxx.ts that works only in NodeJS.
2. Export files in `index.ts` and `index.browser.ts`, they are entries for NodeJS and browser.
3. Keep the exported functions and public ones of class consistent in 2 files.

For example:

```typescript
// onBehalfOfUserCredential.browser.ts
export class OnBehalfOfUserCredential implements TokenCredential {
...
  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    // browser version implementation.
  }
}

// onBehalfOfUserCredential.ts
export class OnBehalfOfUserCredential implements TokenCredential {
...
  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    // nodejs version implementation.
  }
}
```

Please check [onBehalfOfUserCredential.browser.ts](src/credential/onBehalfOfUserCredential.browser.ts) and [onBehalfOfUserCredential.ts](src/credential/onBehalfOfUserCredential.ts) to see the details.
