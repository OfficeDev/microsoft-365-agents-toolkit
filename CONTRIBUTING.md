# Contributing

This project welcomes contributions and suggestions. Most contributions require you to
agree to a Contributor License Agreement (CLA) declaring that you have the right to,
and actually do, grant us the rights to use your contribution. For details, visit
https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need
to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the
instructions provided by the bot. You will only need to do this once across all repositories using our CLA.

# Engineering Guidence
The best ways to contribute to our code is to try things out first, then file the issues you may encountered. Joining the design conversation and make pull request follow our instructions.

For contributing to different packages, find the instructions below:
- [Contributing to VS Code Extension](https://github.com/OfficeDev/TeamsFx/blob/main/packages/vscode-extension/CONTRIBUTING.md)
- [Contributing to SDK](https://github.com/OfficeDev/TeamsFx/blob/main/packages/sdk/CONTRIBUTING.md)
- [Contributing to CLI](https://github.com/OfficeDev/TeamsFx/blob/main/packages/cli/CONTRIBUTING.md)
- [Contributing to API](https://github.com/OfficeDev/TeamsFx/tree/main/packages/api/CONTRIBUTING.md)
- [Contributing to TeamsFx Core](https://github.com/OfficeDev/TeamsFx/tree/main/packages/fx-core/CONTRIBUTING.md)
- [Contributing to Function Extension](https://github.com/OfficeDev/TeamsFx/blob/main/packages/function-extension/CONTRIBUTING.md)
- [Contributing to Simple Auth](https://github.com/OfficeDev/TeamsFx/blob/main/packages/simpleauth/CONTRIBUTING.md)
- [Contributing to .NET SDK](https://github.com/OfficeDev/TeamsFx/blob/main/packages/dotnet-sdk/CONTRIBUTING.md)
- [Contributing to Scaffolding Template](https://github.com/OfficeDev/TeamsFx/blob/main/templates/CONTRIBUTING.md)

## Local development certificates

The engine uses `@peculiar/x509` with Node's native WebCrypto to generate
localhost certificates. The 2.x release line preserves compatibility with the
toolkit's CommonJS builds; import the existing `reflect-metadata` dependency
before the X.509 library. Certificate paths, RSA private-key PEM format, and
Windows/macOS trust-store commands remain unchanged.

Run the certificate and prerequisite regression tests from `packages/fx-core`:

```powershell
pnpm exec vitest run --config vitest.config.ts tests\component\local\localCertificateManager.test.ts tests\component\driver\devTool\installDriver.test.ts
```

For a manual check, debug a Basic Tab locally, approve certificate installation
when prompted, and verify that `https://localhost:3978/tabs/home` loads. Repeat
with an existing valid certificate to confirm it is reused rather than replaced.
Bot-only templates using an HTTP dev-tunnel target do not need this certificate.

# Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)
or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
