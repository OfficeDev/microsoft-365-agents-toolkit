// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.


declare module "vscode" {

	// https://github.com/microsoft/vscode/issues/206587

	export interface AuthenticationGetSessionPresentationOptions {
		/**
		 * An optional Uri to open in the browser to learn more about this authentication request.
		 */
		learnMore?: Uri;
	}
}
