// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.


// https://github.com/microsoft/vscode/issues/229431

declare module "vscode" {

	export namespace window {
		/**
		 * Retrieves the native window handle of the current active window.
		 * This will be updated when the active window changes.
		 */
		export const nativeHandle: Uint8Array | undefined;
	}
}
