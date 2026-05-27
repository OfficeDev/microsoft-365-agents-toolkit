# Product Artifact Viewer

This directory contains the human-facing Markdown/Mermaid preview tool used by `docs/01-product/scenarios/index.html` files.

The viewer is not a product artifact and is not a source of truth. It renders AI-facing `docs/01-product/**/*.md` and `docs/01-product/**/*.mmd` sources for human review only.

Use [`index.html`](index.html) with a product-root-relative `file` query parameter, for example:

```text
index.html?file=ux/README.md
```

Serve the docs directory over local HTTP when previewing, because browser fetch APIs usually cannot read neighboring files from `file://` pages.