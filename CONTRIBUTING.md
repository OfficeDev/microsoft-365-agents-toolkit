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

## Dependency security backports

When a fixed dependency release is unavailable from the required package feed,
use a reviewed PNPM patch rather than suppressing an advisory or changing the
registry. The temporary [node-forge 1.4.0 patch](patches/node-forge@1.4.0.patch)
backports these upstream changes:

- RSA DigestAlgorithm validation: [digitalbazaar/forge#1152](https://github.com/digitalbazaar/forge/pull/1152), revision `ceba34402e329f0365134f23fe19898756527d65` (not merged upstream).
- PEM decoding: [`412345db`](https://github.com/digitalbazaar/forge/commit/412345db583cd28a9d23636d0d93e9f751b935f3).
- PKCS#12 iteration validation: [`d9584aa5`](https://github.com/digitalbazaar/forge/commit/d9584aa56824972858366d5187b3d29e46bf3656).
- TLS record processing: [`817e2ce0`](https://github.com/digitalbazaar/forge/commit/817e2ce0d0f6d8a25bc9c2670615c38a9ffdc51c).
- Fragmented TLS handshakes: [`700d977d`](https://github.com/digitalbazaar/forge/commit/700d977d39f9f8f896c55b0763eebb940a2c087e).
- Certificate extension OIDs: [`58c06bc2`](https://github.com/digitalbazaar/forge/commit/58c06bc251856b409c30ecb35a93f5f6c9bbc84a).

The adapted upstream regression vectors retain the
[node-forge BSD license](patches/node-forge.LICENSE).

Run `pnpm run test:security-backports` after installation. CI verifies the patch
declarations, content hashes, locked dependency references, installed files, and
bounded regression tests for all six advisories. Keep raw audit findings visible:
the installed version remains 1.4.0 and is locally backported, not an
upstream-fixed release.

Replace the five released fixes when a verified fixed version is available in
the feed, but retain the RSA fix until it is also fixed upstream. These workspace
patches protect builds from this checkout; they do not propagate through the
published `@microsoft/teamsfx-core` dependency metadata. Downstream consumers must
upgrade to an upstream-fixed version or apply an equivalent patch themselves.

# Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)
or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
