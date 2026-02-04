# Required Build-time Extensions

Install extensions during Docker build using either extension IDs or `.vsix` files.

**Important**: If any extension in this directory fails to install, the Docker build will fail.

## Method 1: Extension IDs (Recommended)

Add extension IDs to `extensions.txt` file:

```txt
# extensions.txt
ms-python.python
ms-vscode.cpptools
GitHub.copilot
```

- One extension ID per line
- Lines starting with `#` are comments
- Extensions are downloaded from the marketplace during build
- Always gets the latest compatible version

## Method 2: VSIX Files

Place `.vsix` extension files in this directory:

```
build-extensions/
├── extensions.txt
├── ms-python.python-2023.12.0.vsix
└── custom-extension-1.0.0.vsix
```

- Download `.vsix` files from the marketplace
- Useful for specific versions or offline builds
- Supports custom/private extensions

## Usage

1. Choose your method (or use both)
2. Add extension IDs to `extensions.txt` and/or place `.vsix` files here
3. Run `docker build`

These extensions will be permanently installed in the Docker image and available immediately when the container starts.
