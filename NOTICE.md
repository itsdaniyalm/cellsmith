# Third-party notices

Cellsmith itself is MIT-licensed (see `LICENSE`). The built add-in (`dist/`) bundles or loads the following third-party material.

## Bundled in the build

### Monaco Editor
- Copyright (c) Microsoft Corporation. Licensed under the MIT License.
- https://github.com/microsoft/monaco-editor
- Monaco's own list of third-party notices ships in `node_modules/monaco-editor/ThirdPartyNotices.txt`.

### Codicons (icon font used by Monaco's UI)
- Copyright (c) Microsoft Corporation. The icon font is licensed under Creative Commons Attribution 4.0 International (CC BY 4.0); the accompanying code is MIT.
- https://github.com/microsoft/vscode-codicons
- Shipped unmodified as `dist/assets/codicon-*.ttf`.

## Loaded at runtime, not redistributed

### Office.js
- Loaded from Microsoft's CDN (`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`) as required for Office add-ins. It is subject to Microsoft's own terms and is not part of this repository.

## Development tools

Vite, Vitest, TypeScript and the Office add-in tooling are used to build and test the project and are not shipped in `dist/`. Each is under its own license; see their packages.

## Trademarks

Microsoft, Excel, Office, and Microsoft 365 are trademarks of the Microsoft group of companies. Cellsmith is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.
