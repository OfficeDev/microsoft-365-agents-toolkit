# Optional Build-time Extensions

Install extensions during Docker build using either extension IDs or `.vsix` files.

**Note**: If any extension in this directory fails to install, the Docker build will continue (won't fail).

## Method 1: Extension IDs (Recommended)

Add extension IDs to `extensions.txt` file:

```txt
# extensions.txt
esbenp.prettier-vscode
formulahendry.auto-rename-tag
```

- One extension ID per line
- Lines starting with `#` are comments
- Extensions are downloaded from the marketplace during build
- Build continues even if installation fails

## Method 2: VSIX Files

Place `.vsix` extension files in this directory:

```
build-extensions-optional/
├── extensions.txt
├── ms-vscode.theme-monokai-0.1.0.vsix
└── custom-theme-1.0.0.vsix
```

- Download `.vsix` files from the marketplace
- Useful for specific versions or offline builds

## Usage

1. Choose your method (or use both)
2. Add extension IDs to `extensions.txt` and/or place `.vsix` files here
3. Run `docker build` or `.\build.ps1`

These extensions will be installed during build if possible, but build failures won't stop the Docker image creation.
