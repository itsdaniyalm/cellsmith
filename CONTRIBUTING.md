# Contributing to Cellsmith

Thanks for helping out. Bug reports, function-catalog fixes and focused pull requests are all welcome.

## Setup

```bash
npm install
npm test            # unit tests (fast, no Excel needed)
npm run typecheck
npm run dev         # preview at http://localhost:3000/taskpane.html (plain HTTP works without certificates)
```

To try changes inside Excel, follow "Running it" in the README (`npm run certs`, then `npm run start:excel`).

## Where things live

- `src/core/` is pure TypeScript with no browser or Excel dependencies. The tokenizer, parser, formatter and linter live here, and nearly every behaviour change should come with a test in `tests/`.
- `src/core/functionData.ts` is the function catalog. Adding or correcting a function is usually a one-line change; the signature notation is explained at the top of that file. Tests check that every entry parses and has sane argument counts.
- `src/editor/` adapts the core to Monaco (completion, hover, rename, quick fixes).
- `src/excel/` talks to Excel. `session.ts` is the sync state machine and is tested with a mock adapter; `office.ts` is the only code that touches Office.js.

## Before opening a pull request

1. `npm test` and `npm run typecheck` pass.
2. For formatter changes, the corpus tests in `tests/formatter.test.ts` must still pass: formatting must never change what a formula means, and must be idempotent.
3. For linter changes, add a case that passes and a case that fails, and prefer warnings over errors unless Excel itself would reject the formula. Errors block auto-apply.
4. Keep pull requests focused; unrelated cleanups make review harder.

## Reporting bugs

Please include your Excel version and platform (Windows, Mac or web), the formula (remove anything sensitive), what you expected, and what happened. A screenshot of the Problems list helps for linter issues.

## Scope

Cellsmith is deliberately deterministic and local: no AI services and no network calls with workbook data. Proposals that change that are unlikely to be accepted.

## License

By contributing you agree that your contributions are licensed under the MIT License.
