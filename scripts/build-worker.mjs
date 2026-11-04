#!/usr/bin/env node
/**
 * Builds `engine-worker.js`: HyperFormula + this engine in one standalone file for `new Worker(url)`.
 *
 *   node scripts/build-worker.mjs <outfile>
 *
 * Websites copy the output into their static folder (e.g. `public/engine-worker.js`). The file
 * starts with a notice naming the license and the exact source (this repo, at `version`).
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outfile = process.argv[2];
if (!outfile) {
  console.error("Usage: node scripts/build-worker.mjs <outfile>");
  process.exit(2);
}

const pkg = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"));

const banner = `/*!
 * ${pkg.name} ${pkg.version}: spreadsheet engine worker.
 * Copyright (C) 2026 the ${pkg.name} authors. Includes HyperFormula ${pkg.dependencies.hyperformula} (Handsoncode, GPLv3).
 * Licensed under the GNU General Public License v3.0 only. You may copy, change and share it under
 * that license; it comes with no warranty. Full license and complete source code:
 * https://github.com/sabatsibadze1/excel-ge-engine (version ${pkg.version})
 */`;

await build({
  entryPoints: [path.join(root, "src/worker/worker.ts")],
  outfile: path.resolve(outfile),
  bundle: true,
  format: "iife",
  platform: "browser",
  target: "es2020",
  minify: true,
  legalComments: "inline",
  banner: { js: banner },
  logLevel: "warning",
});
console.log(`engine-worker: wrote ${outfile}`);
