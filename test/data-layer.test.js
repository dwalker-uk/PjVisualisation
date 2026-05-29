/*
 * Unit tests for the pure data layer inside roadmap.html.
 *
 * The data-layer functions are DOM/SheetJS-free and live between the
 * "PURE-DATA-LAYER START/END" markers. We extract that section, evaluate it
 * in an isolated VM context, and exercise it with hand-built arrays-of-arrays
 * (the same shape SheetJS produces from the workbook).
 *
 * Run: node test/data-layer.test.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const html = fs.readFileSync(path.join(__dirname, '..', 'roadmap.html'), 'utf8');
const startMarker = html.indexOf('PURE-DATA-LAYER START');
const endMarker = html.indexOf('PURE-DATA-LAYER END');
assert(startMarker !== -1 && endMarker !== -1, 'could not find PURE-DATA-LAYER markers');
// Slice the JS between the two comment blocks (skip past the opening comment's */).
const codeStart = html.indexOf('*/', startMarker) + 2;
const codeEnd = html.lastIndexOf('/*', endMarker);
const code = html.slice(codeStart, codeEnd);

const sandbox = { module: { exports: {} }, console: console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

const { assembleModel, parseExcelDate, colourFor, extractTable } = sandbox;
assert(typeof assembleModel === 'function', 'assembleModel not exported into sandbox');

let passed = 0;
function check(name, fn) {
  try { fn(); passed++; console.log('  ok  - ' + name); }
  catch (e) { console.error('  FAIL- ' + name + '\n        ' + e.message); process.exitCode = 1; }
}

/* ---- Hand-built sheet AOA inputs ---- */

const loesThemes = [
  ['LOE ID', 'Title'],
  ['L1', 'Delivery'],
  ['L2', 'Enablement'],
  [],
  ['LOE ID', 'Theme ID', 'Title', 'Description'],
  ['L1', 'T1', 'Customer Service', 'Improve service'],
  ['L2', 'T2', 'Platform', null],
  [null, 'T-GHOST', null, null], // Excel artefact: auto Theme ID but no parent LOE ID
  []
];

const SERIAL = 45292; // 2024-01-01 in Excel serial (1900 system)

const activities = [
  ['Theme ID', 'Activity ID', 'UniqueAct ID', 'Level', 'Title', 'Description', 'Start Date', 'End Date', 'Resourcing'],
  ['T1', 'A', 'U-A', 1, 'Top activity', 'desc', SERIAL, SERIAL + 200, 'Funded'],
  ['T1', 'A-1', 'U-A1', 2, 'Child one', '', SERIAL + 10, SERIAL + 80, 'At risk'],
  ['T1', 'A-1-x', 'U-A1x', 3, 'Grandchild', '', SERIAL + 11, SERIAL + 40, 'Funded'],
  ['T2', 'B', 'U-B', 1, 'Platform work', '', new Date(2026, 0, 1), new Date(2026, 5, 1), 'Unfunded'],
  ['T1', 'C-orphan', 'U-Corph', 2, 'Orphan child', '', SERIAL, SERIAL + 100, 'Funded'],
  [null, null, 'U-GHOST', null, null, null, null, null, null] // artefact: auto UniqueAct ID, no parent Theme ID
];

const milestones = [
  ['Activity ID', 'UniqueMSt ID', 'Milestone Title', 'Description', 'Date', 'Owner', 'Delivery Confidence'],
  ['U-A', 'M1', 'First milestone', '', SERIAL + 20, 'Alice', 'High'],
  ['U-A1', 'M2', 'Child milestone', '', SERIAL + 30, 'Bob', 'Low'],
  ['U-B', 'M3', 'Platform milestone', '', new Date(2026, 2, 1), 'Carol', 'Medium'],
  [null, 'M-GHOST', null, null, null, null, null] // artefact: auto UniqueMSt ID, no parent Activity ID
];

const benefits = [
  ['Milestone ID', 'Benefit Title', 'Category', 'Beneficiary', 'Impact'],
  ['M1', 'Faster service', 'Efficiency', 'Citizens', 'High']
];

const risks = [
  ['Milestone ID', 'Risk Title', 'Risk Owner', 'RAG'],
  ['M2', 'Resource gap', 'Dave', 'Red']
];

// Lookups table anchored at F1 (column index 5). Columns 0..4 are empty.
const E = [null, null, null, null, null];
const lookups = [
  E.concat(['Activities', 'Milestones', 'Risks']),
  E.concat(['Resourcing', 'Delivery Confidence', 'RAG']),
  E.concat(['green;amber;grey', 'green;amber;red', 'red;amber;green']),
  E.concat(['Funded', 'High', 'Red']),
  E.concat(['At risk', 'Medium', 'Amber']),
  E.concat(['Unfunded', 'Low', 'Green'])
];

const model = assembleModel({ loesThemes, activities, milestones, benefits, risks, lookups });

/* ---- Tests ---- */

check('parses LOEs and Themes from the shared sheet', () => {
  assert.strictEqual(model.loes.length, 2);
  assert.strictEqual(model.themes.length, 2);
  assert.strictEqual(model.loeById['L1'].title, 'Delivery');
  assert.strictEqual(model.themeById['T1'].loeId, 'L1');
  assert.strictEqual(model.themeById['T1'].description, 'Improve service');
});

check('reads all activities, keyed by UniqueAct ID', () => {
  assert.strictEqual(model.activities.length, 5);
  assert(model.activityByUid['U-A']);
  assert.strictEqual(model.activityByUid['U-A'].resourcing, 'Funded');
});

check('drops Excel trailing-artefact rows via the parent-id column', () => {
  assert.strictEqual(model.themes.length, 2, 'ghost theme dropped');
  assert(!model.themeById['T-GHOST']);
  assert(!model.activityByUid['U-GHOST'], 'ghost activity dropped');
  assert.strictEqual(model.milestones.length, 3, 'ghost milestone dropped');
  assert(!model.milestoneByUid['M-GHOST']);
});

check('builds dash-delimited hierarchy', () => {
  assert.strictEqual(model.activityByUid['U-A'].depth, 1);
  assert.strictEqual(model.activityByUid['U-A1'].depth, 2);
  assert.strictEqual(model.activityByUid['U-A1'].parentUid, 'U-A');
  assert.strictEqual(model.activityByUid['U-A'].childUids.length, 1);
  assert.strictEqual(model.activityByUid['U-A'].childUids[0], 'U-A1');
  assert.strictEqual(model.activityByUid['U-A1x'].depth, 3);
  assert.strictEqual(model.activityByUid['U-A1x'].parentUid, 'U-A1');
});

check('promotes orphan activity to top level with a warning', () => {
  assert.strictEqual(model.activityByUid['U-Corph'].depth, 1);
  assert(model.warnings.some(w => w.id === 'U-Corph' && /no parent/.test(w.reason)));
});

check('links milestones, benefits and risks by milestone uid', () => {
  assert.strictEqual(model.milestones.length, 3);
  assert.strictEqual((model.milestonesByActivity['U-A'] || []).length, 1);
  assert.strictEqual((model.benefitsByMilestone['M1'] || []).length, 1);
  assert.strictEqual((model.risksByMilestone['M2'] || []).length, 1);
});

check('ignores blank Benefit/Risk rows that only carry a formula Milestone ID', () => {
  // Trailing rows where the Milestone ID is an auto-filled formula but every
  // content column is empty must not appear as benefits/risks.
  const m2 = assembleModel({
    loesThemes, activities, milestones, lookups,
    benefits: benefits.concat([['M1', null, null, null, null]]),
    risks: risks.concat([['M2', '', '', '']])
  });
  assert.strictEqual(m2.benefits.length, 1, 'blank benefit row dropped');
  assert.strictEqual(m2.risks.length, 1, 'blank risk row dropped');
  assert.strictEqual((m2.benefitsByMilestone['M1'] || []).length, 1);
  assert.strictEqual((m2.risksByMilestone['M2'] || []).length, 1);
});

check('resolves colours from Lookups, brand palette + literal fallback', () => {
  assert.strictEqual(colourFor(model.lookups, 'Activities', 'Resourcing', 'Funded'), '#79C68A');   // green
  assert.strictEqual(colourFor(model.lookups, 'Activities', 'Resourcing', 'Unfunded'), 'grey');     // literal CSS
  assert.strictEqual(colourFor(model.lookups, 'Milestones', 'Delivery Confidence', 'Low'), '#D41F41'); // red
  assert.strictEqual(colourFor(model.lookups, 'Risks', 'RAG', 'Amber'), '#E08D48');                 // orange/amber
});

check('colour lookup is case-insensitive and falls back', () => {
  assert.strictEqual(colourFor(model.lookups, 'Activities', 'Resourcing', 'funded'), '#79C68A');
  assert.strictEqual(colourFor(model.lookups, 'Activities', 'Resourcing', 'Nonsense', '#000000'), '#000000');
});

check('parses Excel serial dates and Date objects', () => {
  const d = parseExcelDate(SERIAL);
  assert.strictEqual(d.getFullYear(), 2024);
  assert.strictEqual(d.getMonth(), 0);
  assert.strictEqual(d.getDate(), 1);
  const iso = parseExcelDate('2025-06-15');
  assert.strictEqual(iso.getFullYear(), 2025);
  assert.strictEqual(iso.getMonth(), 5);
  const uk = parseExcelDate('15/06/2025');
  assert.strictEqual(uk.getMonth(), 5);
  assert.strictEqual(uk.getDate(), 15);
});

check('exposes distinct Risk RAG / Benefit Category values for filters', () => {
  assert.deepStrictEqual(Array.from(model.riskRagValues), ['Red']);
  assert.deepStrictEqual(Array.from(model.benefitCategoryValues), ['Efficiency']);
});

check('computes the data date extent', () => {
  // Dates are created inside the VM realm, so check by duck-typing not instanceof.
  assert(model.dataDateMin && typeof model.dataDateMin.getFullYear === 'function');
  assert(model.dataDateMax && typeof model.dataDateMax.getFullYear === 'function');
  assert(model.dataDateMin <= model.dataDateMax);
});

check('extractTable stops at the first blank row', () => {
  const t = extractTable(loesThemes, { id: ['loe id'], title: ['title'] }, ['id', 'title'], { rejectHeaders: ['theme id'] });
  assert.strictEqual(t.records.length, 2);
  assert.strictEqual(t.records[0].id, 'L1');
});

console.log('\n' + passed + ' checks passed.');
