# Aixa PPTX Viewer

A React-based PowerPoint viewer and editor maintained by Aixa Ltd. It supports
rendering, editing, presenting, and exporting PPTX presentations in the browser.

- [npm package](https://www.npmjs.com/package/@aixa-transformation/pptx-viewer)
- [React package documentation](packages/react/README.md)

## Attribution

This project is a fork of
[ChristopherVR/pptx-viewer](https://github.com/ChristopherVR/pptx-viewer).
The original project is Copyright © 2025-present pptx-viewer contributors and
is licensed under the Apache License 2.0.

This fork is maintained by Aixa Ltd and contains modifications made by the Aixa
Transformation team. See the `LICENSE` and `NOTICE` files included with each
distributed package for licensing and attribution details.

## Packages

- `packages/react`: React viewer and editor component
- `packages/core`: Framework-independent PPTX engine
- `packages/shared`: Internal logic shared by viewer packages

## Build a local package archive

With dependencies installed, run from the repository root:

```sh
npm run pack:tgz
```

This builds the core, shared, and React packages, then creates
`aixa-transformation-pptx-viewer-<version>.tgz` in the repository root. The version
comes from `packages/react/package.json`; it is not incremented automatically.
Rebuilding the same version replaces its existing archive. Nothing is published
to npm, and a failed build stops packaging.

From `packages/react`, `npm run pack` runs the same workflow. The final pack step
skips lifecycle scripts because the build has already run, avoiding an extra
`prepack` build.

## License

Licensed under the Apache License 2.0. See the package-level `LICENSE` and
`NOTICE` files for details.
