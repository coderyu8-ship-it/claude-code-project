#!/usr/bin/env node
/**
 * Rank source files by coverage, lowest first, so weak spots are obvious.
 *
 * Reads coverage/coverage-summary.json, written by `npm run test:coverage`.
 *
 *   node scripts/coverage-report.mjs            # report only
 *   node scripts/coverage-report.mjs --min 80   # also exit 1 if a file is below 80%
 */

import { readFileSync } from 'node:fs';
import { relative, resolve } from 'node:path';
import process from 'node:process';

const SUMMARY_PATH = resolve('coverage/coverage-summary.json');
const METRICS = ['statements', 'branches', 'functions', 'lines'];

function parseArgs(argv) {
  const minIndex = argv.indexOf('--min');
  if (minIndex === -1) {
    return { min: null };
  }

  const raw = argv[minIndex + 1];
  const min = Number(raw);
  if (!Number.isFinite(min) || min < 0 || min > 100) {
    console.error(`--min expects a number between 0 and 100, received ${raw ?? '(nothing)'}`);
    process.exit(2);
  }
  return { min };
}

function readSummary() {
  try {
    return JSON.parse(readFileSync(SUMMARY_PATH, 'utf8'));
  } catch (error) {
    if (error.code === 'ENOENT') {
      console.error('coverage/coverage-summary.json not found. Run `npm run test:coverage` first.');
      process.exit(2);
    }
    throw error;
  }
}

function pct(entry, metric) {
  const value = entry?.[metric];
  // istanbul reports 100% for a file with nothing of that kind to cover
  // (e.g. no branches at all); treat that as covered rather than as a hole.
  if (value === undefined || value.total === 0) {
    return 100;
  }
  return value.pct;
}

const { min } = parseArgs(process.argv.slice(2));
const summary = readSummary();
const { total, ...files } = summary;

const rows = Object.entries(files)
  .map(([path, entry]) => ({
    path: relative(process.cwd(), path),
    statements: pct(entry, 'statements'),
    branches: pct(entry, 'branches'),
    functions: pct(entry, 'functions'),
    lines: pct(entry, 'lines'),
    uncoveredLines: entry.lines?.total - entry.lines?.covered,
  }))
  .map((row) => ({ ...row, worst: Math.min(...METRICS.map((metric) => row[metric])) }))
  .sort((a, b) => a.worst - b.worst || a.path.localeCompare(b.path));

const fmt = (value) => `${value.toFixed(2).padStart(6)}%`;
const width = Math.max(20, ...rows.map((row) => row.path.length));

console.log('\nCoverage by file (weakest first)\n');
console.log(
  `${'FILE'.padEnd(width)}  ${'STMTS'.padStart(7)} ${'BRANCH'.padStart(7)} ${'FUNCS'.padStart(7)} ${'LINES'.padStart(7)}  UNCOVERED`,
);
console.log('-'.repeat(width + 46));

for (const row of rows) {
  const flag = min !== null && row.worst < min ? '  <-- below threshold' : '';
  console.log(
    `${row.path.padEnd(width)}  ${fmt(row.statements)} ${fmt(row.branches)} ${fmt(row.functions)} ${fmt(row.lines)}  ${String(row.uncoveredLines).padStart(9)}${flag}`,
  );
}

console.log('-'.repeat(width + 46));
console.log(
  `${'TOTAL'.padEnd(width)}  ${fmt(pct(total, 'statements'))} ${fmt(pct(total, 'branches'))} ${fmt(pct(total, 'functions'))} ${fmt(pct(total, 'lines'))}\n`,
);

if (min !== null) {
  const failing = rows.filter((row) => row.worst < min);
  if (failing.length > 0) {
    console.error(`${failing.length} file(s) below the ${min}% floor:`);
    for (const row of failing) {
      console.error(`  ${row.path} (worst metric: ${row.worst.toFixed(2)}%)`);
    }
    process.exit(1);
  }
  console.log(`All files are at or above the ${min}% floor.`);
}
