# CLAUDE.md

Public GPLv3 engine of a Georgian Excel practice simulator. Read `README.md`, then `src/engine/README.md`.

- The website using this engine is private. It talks to `engine-worker.js` only through the JSON messages in
  `src/worker/protocol.ts`, and imports nothing from here except `excel-ge-engine/pure`. Keep it that way:
  nothing reachable from `src/pure.ts` may import `hyperformula` (`src/pure.test.ts` checks).
- Never add lessons, facts, test banks, prices or business documents here. Only the small test fixtures in
  `src/checker/fixtures/` belong here.
- HyperFormula 3.5.0, pinned exactly, `licenseKey: 'gpl-v3'`, always via `createEngine()`, never `updateConfig()`.
- Grading is swap-the-inputs, never regex. Format conversion via `enToKa`/`kaToEn` only.
- Before finishing: `npm run typecheck`, `npm test`, `npm run build`.
