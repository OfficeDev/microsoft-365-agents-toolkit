# M365 Spec Parser

The M365 Spec Parser package is designed to parse OpenAPI specification files to generate resources for M365 applications.

## Requirements

For Node.js, requires version 20.19.0 or later in the 20.x series, or version 22.12.0 or later, matching the OpenAPI parser dependency's runtime requirements.

## Sample Usage

### For NodeJS environment
```typescript
import { SpecParser, ParseOptions } from "@microsoft/m365-spec-parser";

// Define parsing options
const option: ParseOptions = {
  allowMissingId: true, // Allow missing IDs in the specification, default true
  allowSwagger: true, // Allow Swagger specifications, default true
  allowAPIKeyAuth: false, // Disallow API key authentication, default false
  allowMultipleParameters: false, // Disallow multiple parameters, default false
  allowOauth2: false, // Disallow OAuth2 authentication, default false
};

// Create a new SpecParser instance with the given options
const parser = new SpecParser("path/to/spec/file", option);

// Validate the specification
const validateResult = await parser.validate();
// If the specification is not valid, log the errors and warnings
if (validateResult.status !== ValidationStatus.Valid) {
  console.log(validateResult.errors);
  console.log(validateResult.warnings);
}

// List the operations in the specification
const listResult = await parser.list();
// Log each operation
for (let i = 0; i < listResult.length; i++) {
  console.log(listResult[i]);
}

// Define a filter for the operations to generate
const filter = ["GET /pet/{id}"];
// Define the paths for the Teams app manifest file, the output specification file, and the output adaptive card folder
const teamsAppManifestFilePath = "path/to/teamsapp/manifest/file";
const outputSpecFilePath = "path/to/output/spec/path";
const outputAdaptiveCardFolder = "adaptivecard/folder";
// Generate the operations
const generateResult = await parser.generate(
  teamsAppManifestFilePath,
  filter,
  outputSpecFilePath,
  outputAdaptiveCardFolder
);

// If not all operations were successfully generated, log the warnings
if (!generateResult.allSuccess) {
  console.log(generateResult.warnings);
}
```

### For browser environment
```typescript
import { SpecParser, ParseOptions } from "@microsoft/m365-spec-parser";

// Define parsing options
const option: ParseOptions = {
  allowMissingId: false, // Allow missing IDs in the specification, default false
  allowSwagger: false, // Allow Swagger specifications, default false
  allowAPIKeyAuth: false, // Disallow API key authentication, default false
  allowMultipleParameters: false, // Disallow multiple parameters, default false
  allowOauth2: false, // Disallow OAuth2 authentication, default false
};

// Create a new SpecParser instance with the given options
const parser = new SpecParser("path/to/spec/file", option);

// Validate the specification
const validateResult = await parser.validate();
// If the specification is not valid, log the errors and warnings
if (validateResult.status !== ValidationStatus.Valid) {
  console.log(validateResult.errors);
  console.log(validateResult.warnings);
}

// List the operations in the specification
const listResult = await parser.listSupportedAPIInfo();
// Log each operation
for (let i = 0; i < listResult.length; i++) {
  console.log(listResult[i]);
}
```

## Data Collection.

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at https://go.microsoft.com/fwlink/?LinkID=824704. You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)

If you are interested in fixing issues and contributing directly to the code base, please see the [Contributing Guide](./CONTRIBUTING.md).

## Reporting Security Issues

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them to the Microsoft Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report).

If you prefer to submit without logging in, send email to [secure@microsoft.com](mailto:secure@microsoft.com). If possible, encrypt your message with our PGP key; please download it from the the [Microsoft Security Response Center PGP Key page](https://www.microsoft.com/en-us/msrc/pgp-key-msrc).

You should receive a response within 24 hours. If for some reason you do not, please follow up via email to ensure we received your original message. Additional information can be found at [microsoft.com/msrc](https://www.microsoft.com/msrc).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE.txt) license.
