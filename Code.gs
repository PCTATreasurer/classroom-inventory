/**
 * Classroom Inventory — Google Sheets + Apps Script
 * --------------------------------------------------
 * Tabs this script builds:
 *   Catalog      – the full district Warehouse Digital Catalog (imported, with pictures)
 *   Inventory    – the items you actually keep (catalog items + your own items)
 *   Scan         – scan barcodes here (USB/Bluetooth scanner or typing)
 *   Count Sheet  – printable / on-screen count form with pictures + condition
 *   Labels       – printable barcode labels
 *   Planning     – how much of each item you need next year
 *   Order List   – generated shortfall list (need − usable on hand)
 *
 * Setup: Extensions ▸ Apps Script ▸ paste this file ▸ Save ▸ reload the sheet ▸
 *        Inventory menu ▸ "1. Set up / import catalog".
 */

const CATALOG_URL =
  'https://docs.google.com/document/d/e/2PACX-1vTQkkm6v1tQ9O-oAWwd5BRcZJrJPRRVydJykFdSkMvt9YddeGt6zpDXISEXXFIIEL1hAsS2Y4piLoE3/pub';

const TYPES = ['Consumable', 'Durable'];
const MODES = ['Scan each', 'Enter qty'];
const CONDITIONS = ['Good', 'Fair', 'Needs repair', 'Replace'];
const BARCODE_FONT = 'Libre Barcode 39 Text'; // Code 39 – every scanner reads it by default

// Column maps (1-based)
const CAT = { track: 1, pic: 2, stock: 3, desc: 4, unit: 5, cost: 6, section: 7, type: 8, img: 9 };
const CAT_HEAD = ['Track?', 'Picture', 'Stock #', 'Description', 'Unit', 'Unit Cost', 'Section', 'Type', 'Image ID'];

const INV = { pic: 1, id: 2, desc: 3, unit: 4, type: 5, mode: 6, loc: 7, onHand: 8, cond: 9,
              replace: 10, counted: 11, cost: 12, source: 13, img: 14, label: 15, notes: 16 };
const INV_HEAD = ['Picture', 'Item ID', 'Description', 'Unit', 'Type', 'Count Mode', 'Location',
                  'On Hand', 'Condition', '# To Replace', 'Last Counted', 'Unit Cost', 'Source',
                  'Image ID', 'Label?', 'Notes'];

const SCAN_HEAD = ['Scan here ▼', 'Qty', 'Item', 'Mode', 'Time', 'Status'];
const PLAN_HEAD = ['Item ID', 'Description', 'Unit', 'Per Student', 'Students', 'Extra', 'Total Need', 'Notes'];
const ORDER_HEAD = ['Stock / Item ID', 'Description', 'Unit', 'Need', 'Usable On Hand', 'Order Qty',
                    'Unit Cost', 'Line Total', 'Source', 'Reason'];

const ROW_H = 64; // picture row height (px)

// ───────────────────────────── Menu ─────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Inventory')
    .addItem('1. Set up / import catalog', 'setup')
    .addItem('Refresh catalog (prices & pictures)', 'importCatalog')
    .addSeparator()
    .addItem('Add checked catalog items → Inventory', 'addTrackedToInventory')
    .addItem('Add my own (non-catalog) item', 'addCustomItem')
    .addSeparator()
    .addItem('Build Count Sheet', 'buildCountSheet')
    .addItem('Apply Count Sheet → Inventory', 'applyCountSheet')
    .addItem('Apply Scans → Inventory', 'applyScans')
    .addItem('Clear Scan tab', 'clearScans')
    .addSeparator()
    .addItem('Build barcode labels', 'buildLabels')
    .addSeparator()
    .addItem('Add Inventory items → Planning', 'syncPlanning')
    .addItem('Build Order List', 'buildOrderList')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, 'Catalog', CAT_HEAD);
  const inv = ensureSheet_(ss, 'Inventory', INV_HEAD);
  ensureSheet_(ss, 'Scan', SCAN_HEAD);
  ensureSheet_(ss, 'Planning', PLAN_HEAD);
  ensureSheet_(ss, 'Order List', ORDER_HEAD);
  ensureSheet_(ss, 'Count Sheet', []);
  ensureSheet_(ss, 'Labels', []);
  importCatalog();
  ss.setActiveSheet(ss.getSheetByName('Catalog'));
  toast_('Done. Tick "Track?" on the items you keep, then Inventory ▸ Add checked catalog items.');
}

// ───────────────────────────── Catalog import ─────────────────────────────

/** Parses the published Google Doc HTML. Returns [section, stock, unit, desc, cost, imageId]. */
function parseCatalog_(html) {
  const ent = s => s.replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
  const txt = s => ent(s.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
  const start = html.indexOf('<table');
  const end = html.indexOf('</table>', start);
  const tbl = html.slice(start, end > 0 ? end : undefined);
  const trs = tbl.split(/<tr\b[^>]*>/).slice(1);
  let sec = '';
  const out = [];
  for (const tr of trs) {
    const tds = tr.split(/<td\b[^>]*>/).slice(1);
    const cells = tds.map(txt);
    if (cells.length < 4) { const j = cells.join(' ').trim(); if (j) sec = j; continue; }
    const stock = cells[0];
    if (!/^\d{4,5}$/.test(stock)) continue; // skips header rows
    let id = '';
    const h = (tds[0].match(/href="([^"]+)"/) || [])[1];
    if (h) {
      const d = decodeURIComponent(ent(h));
      const m = d.match(/\/d\/([\w-]{20,})/) || d.match(/[?&]id=([\w-]{20,})/);
      if (m) id = m[1];
    }
    out.push([sec, stock.padStart(5, '0'), cells[1], cells[2], cells[3], id]);
  }
  return out;
}

/** "$0..23" → 0.23, "$1,204.00" → 1204 */
function parseCost_(s) {
  const n = parseFloat(String(s).replace(/[$,\s]/g, '').replace(/\.{2,}/g, '.'));
  return isNaN(n) ? '' : n;
}

function importCatalog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureSheet_(ss, 'Catalog', CAT_HEAD);
  const html = UrlFetchApp.fetch(CATALOG_URL, { muteHttpExceptions: true }).getContentText();
  const items = parseCatalog_(html);
  if (!items.length) throw new Error('No catalog rows found — has the district catalog layout changed?');

  // keep your Track?/Type choices across refreshes
  const keep = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, CAT_HEAD.length).getValues().forEach(r => {
      keep[r[CAT.stock - 1]] = { track: r[CAT.track - 1] === true, type: r[CAT.type - 1] };
    });
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
  }

  const rows = items.map(([sec, stock, unit, desc, cost, img]) => {
    const k = keep[stock] || {};
    return [k.track || false, '', stock, desc, unit, parseCost_(cost), sec, k.type || 'Consumable', img];
  });
  const n = rows.length;
  sh.getRange(2, CAT.stock, n, 1).setNumberFormat('@');
  sh.getRange(2, 1, n, CAT_HEAD.length).setValues(rows);
  sh.getRange(2, CAT.pic, n, 1).setFormulaR1C1(picFormula_(CAT.img - CAT.pic));
  sh.getRange(2, CAT.track, n, 1).insertCheckboxes();
  sh.getRange(2, CAT.type, n, 1).setDataValidation(list_(TYPES));
  sh.getRange(2, CAT.cost, n, 1).setNumberFormat('$#,##0.00');
  sh.setRowHeights(2, n, ROW_H);
  sh.setColumnWidth(CAT.pic, 90).setColumnWidth(CAT.desc, 360).setColumnWidth(CAT.section, 130);
  sh.hideColumns(CAT.img);
  sh.setFrozenRows(1);
  if (!sh.getFilter()) sh.getRange(1, 1, n + 1, CAT_HEAD.length).createFilter();

  // push price/picture updates into Inventory
  const inv = ss.getSheetByName('Inventory');
  if (inv && inv.getLastRow() > 1) {
    const byStock = {};
    rows.forEach(r => byStock[r[CAT.stock - 1]] = r);
    const rng = inv.getRange(2, 1, inv.getLastRow() - 1, INV_HEAD.length);
    const vals = rng.getValues();
    vals.forEach(v => {
      const c = byStock[String(v[INV.id - 1])];
      if (c) { v[INV.cost - 1] = c[CAT.cost - 1]; v[INV.img - 1] = c[CAT.img - 1]; }
    });
    inv.getRange(2, INV.cost, vals.length, 1).setValues(vals.map(v => [v[INV.cost - 1]]));
    inv.getRange(2, INV.img, vals.length, 1).setValues(vals.map(v => [v[INV.img - 1]]));
  }
  toast_(`Catalog imported: ${n} items (${rows.filter(r => r[CAT.img - 1]).length} with pictures).`);
}

// ───────────────────────────── Inventory ─────────────────────────────

function addTrackedToInventory() {
  const ss = SpreadsheetApp.getActive();
  const cat = ss.getSheetByName('Catalog');
  const inv = ensureSheet_(ss, 'Inventory', INV_HEAD);
  const have = new Set(colValues_(inv, INV.id).map(String));
  const add = cat.getRange(2, 1, cat.getLastRow() - 1, CAT_HEAD.length).getValues()
    .filter(r => r[CAT.track - 1] === true && !have.has(String(r[CAT.stock - 1])))
    .map(r => invRow_({
      id: r[CAT.stock - 1], desc: r[CAT.desc - 1], unit: r[CAT.unit - 1], type: r[CAT.type - 1],
      cost: r[CAT.cost - 1], source: 'District warehouse', img: r[CAT.img - 1]
    }));
  if (!add.length) return toast_('Nothing new — tick "Track?" on Catalog rows first.');
  appendInventory_(inv, add);
  toast_(`Added ${add.length} item(s) to Inventory. Set Location and Count Mode there.`);
}

function addCustomItem() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Add my own item', 'Description (e.g. "Duct tape, 1.88 in roll")', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText().trim()) return;
  const inv = ensureSheet_(SpreadsheetApp.getActive(), 'Inventory', INV_HEAD);
  const ids = colValues_(inv, INV.id).map(String).filter(s => /^E-\d+$/.test(s)).map(s => +s.slice(2));
  const id = 'E-' + String((ids.length ? Math.max(...ids) : 0) + 1).padStart(4, '0');
  appendInventory_(inv, [invRow_({ id, desc: r.getResponseText().trim(), unit: 'Each', type: 'Consumable', source: 'Other' })]);
  toast_(`Added ${id}. Paste a picture link's file ID in "Image ID" if you have one.`);
}

function invRow_(o) {
  const row = new Array(INV_HEAD.length).fill('');
  row[INV.id - 1] = o.id; row[INV.desc - 1] = o.desc; row[INV.unit - 1] = o.unit;
  row[INV.type - 1] = o.type || 'Consumable'; row[INV.mode - 1] = 'Scan each';
  row[INV.cost - 1] = o.cost === undefined ? '' : o.cost; row[INV.source - 1] = o.source || '';
  row[INV.img - 1] = o.img || ''; row[INV.label - 1] = false;
  return row;
}

function appendInventory_(inv, rows) {
  const start = inv.getLastRow() + 1;
  inv.getRange(start, INV.id, rows.length, 1).setNumberFormat('@');
  inv.getRange(start, 1, rows.length, INV_HEAD.length).setValues(rows);
  formatInventory_(inv);
}

function formatInventory_(inv) {
  if (inv.getLastRow() < 2) return;
  const n = inv.getLastRow() - 1;
  inv.getRange(2, INV.pic, n, 1).setFormulaR1C1(picFormula_(INV.img - INV.pic));
  inv.getRange(2, INV.type, n, 1).setDataValidation(list_(TYPES));
  inv.getRange(2, INV.mode, n, 1).setDataValidation(list_(MODES));
  inv.getRange(2, INV.cond, n, 1).setDataValidation(list_(CONDITIONS));
  inv.getRange(2, INV.label, n, 1).insertCheckboxes();
  inv.getRange(2, INV.cost, n, 1).setNumberFormat('$#,##0.00');
  inv.getRange(2, INV.counted, n, 1).setNumberFormat('m/d/yyyy');
  inv.setRowHeights(2, n, ROW_H);
  inv.setColumnWidth(INV.pic, 90).setColumnWidth(INV.desc, 320);
  inv.setFrozenRows(1);
  inv.setFrozenColumns(INV.id);
}

// ───────────────────────────── Scanning ─────────────────────────────
// A USB/Bluetooth scanner "types" the code and presses Enter, so each scan lands
// on a new row in column A. "Scan each" items get Qty 1 per scan; "Enter qty"
// items (e.g. tape) turn the Qty cell yellow — type the number there.

function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Scan' || e.range.getRow() < 2) return;
  const col = e.range.getColumn();
  if (col === 2) { e.range.setBackground(null); return; }
  if (col !== 1 || e.range.getNumRows() > 1) return;

  const raw = String(e.range.getValue() || '').trim().replace(/^\*|\*$/g, '').toUpperCase();
  if (!raw) return;
  const code = /^\d{4}$/.test(raw) ? raw.padStart(5, '0') : raw;
  e.range.setValue(code);
  const out = e.range.offset(0, 1, 1, 5);
  const inv = e.source.getSheetByName('Inventory');
  const vals = inv.getRange(2, 1, Math.max(inv.getLastRow() - 1, 1), INV_HEAD.length).getValues();
  const hit = vals.find(v => String(v[INV.id - 1]).toUpperCase() === code);
  if (!hit) {
    out.setValues([['', 'NOT IN INVENTORY', '', new Date(), 'Error']]).setBackground('#f4cccc');
    return;
  }
  const mode = hit[INV.mode - 1] || 'Scan each';
  const qtyCell = e.range.offset(0, 1);
  const qty = mode === 'Enter qty' ? qtyCell.getValue() : (qtyCell.getValue() || 1);
  out.setValues([[qty, hit[INV.desc - 1], mode, new Date(), '']]).setBackground(null);
  if (mode === 'Enter qty' && qty === '') qtyCell.setBackground('#fff2cc').activate();
}

/** Totals all un-applied scans per item and writes them as the new On Hand. */
function applyScans() {
  const ss = SpreadsheetApp.getActive();
  const sc = ss.getSheetByName('Scan');
  if (sc.getLastRow() < 2) return toast_('No scans yet.');
  const rows = sc.getRange(2, 1, sc.getLastRow() - 1, SCAN_HEAD.length).getValues();
  const totals = {};
  const missing = [];
  rows.forEach((r, i) => {
    if (!r[0] || r[5] === 'Applied' || r[5] === 'Error') return;
    if (r[1] === '' || isNaN(r[1])) { missing.push(i + 2); return; }
    totals[r[0]] = (totals[r[0]] || 0) + Number(r[1]);
  });
  if (missing.length) return toast_(`Fill in Qty on Scan row(s) ${missing.join(', ')} first.`);
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Apply scans',
    `Set On Hand for ${Object.keys(totals).length} item(s) to the scanned totals?\n` +
    '(Items you did not scan are left unchanged.)', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  setOnHand_(totals, {});
  rows.forEach((r, i) => { if (r[0] && r[5] === '') sc.getRange(i + 2, 6).setValue('Applied'); });
  toast_('Inventory updated from scans.');
}

function clearScans() {
  const sc = SpreadsheetApp.getActive().getSheetByName('Scan');
  if (sc.getLastRow() > 1) sc.getRange(2, 1, sc.getLastRow() - 1, SCAN_HEAD.length).clearContent().setBackground(null);
  sc.getRange('A2').activate();
}

function setOnHand_(totals, extras) {
  const inv = SpreadsheetApp.getActive().getSheetByName('Inventory');
  const rng = inv.getRange(2, 1, inv.getLastRow() - 1, INV_HEAD.length);
  const vals = rng.getValues();
  const today = new Date();
  vals.forEach(v => {
    const id = String(v[INV.id - 1]);
    if (id in totals) { v[INV.onHand - 1] = totals[id]; v[INV.counted - 1] = today; }
    if (extras[id]) {
      if (extras[id].cond) v[INV.cond - 1] = extras[id].cond;
      if (extras[id].replace !== '') v[INV.replace - 1] = extras[id].replace;
    }
  });
  [INV.onHand, INV.cond, INV.replace, INV.counted].forEach(c =>
    inv.getRange(2, c, vals.length, 1).setValues(vals.map(v => [v[c - 1]])));
}

// ───────────────────────────── Count Sheet ─────────────────────────────

const CS_HEAD = ['Picture', 'Location', 'Item ID', 'Description', 'Unit', 'Count', 'Condition', '# To Replace', 'Image ID'];

function buildCountSheet() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventory');
  if (inv.getLastRow() < 2) return toast_('Inventory is empty.');
  const items = inv.getRange(2, 1, inv.getLastRow() - 1, INV_HEAD.length).getValues()
    .filter(v => v[INV.id - 1])
    .sort((a, b) => String(a[INV.loc - 1]).localeCompare(String(b[INV.loc - 1])) ||
                    String(a[INV.desc - 1]).localeCompare(String(b[INV.desc - 1])));
  const cs = ensureSheet_(ss, 'Count Sheet', []);
  cs.getRange(1, 1, cs.getMaxRows(), cs.getMaxColumns()).clearDataValidations();
  cs.clear();
  cs.getRange(1, 1).setValue('Inventory Count — ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMM d, yyyy'))
    .setFontSize(14).setFontWeight('bold');
  cs.getRange(2, 1, 1, CS_HEAD.length).setValues([CS_HEAD]).setFontWeight('bold').setBackground('#d9ead3');
  const rows = items.map(v => {
    const durable = v[INV.type - 1] === 'Durable';
    return ['', v[INV.loc - 1], v[INV.id - 1], v[INV.desc - 1], v[INV.unit - 1], '',
            '', durable ? '' : '—', v[INV.img - 1]];
  });
  const n = rows.length;
  cs.getRange(3, 3, n, 1).setNumberFormat('@');
  cs.getRange(3, 1, n, CS_HEAD.length).setValues(rows);
  cs.getRange(3, 1, n, 1).setFormulaR1C1(picFormula_(8));
  items.forEach((v, i) => {
    if (v[INV.type - 1] === 'Durable') cs.getRange(3 + i, 7).setDataValidation(list_(CONDITIONS));
    else cs.getRange(3 + i, 7).setValue('—');
  });
  cs.setRowHeights(3, n, 80);
  cs.setColumnWidth(1, 100).setColumnWidth(2, 110).setColumnWidth(3, 70).setColumnWidth(4, 300)
    .setColumnWidth(5, 70).setColumnWidth(6, 80).setColumnWidth(7, 110).setColumnWidth(8, 90);
  cs.getRange(2, 1, n + 1, 8).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
  cs.getRange(3, 6, n, 1).setBackground('#fff2cc').setFontSize(14).setHorizontalAlignment('center');
  cs.hideColumns(9);
  cs.setFrozenRows(2);
  ss.setActiveSheet(cs);
  toast_('Count Sheet ready — print it (File ▸ Print, fit to width) or type counts in the yellow column.');
}

function applyCountSheet() {
  const cs = SpreadsheetApp.getActive().getSheetByName('Count Sheet');
  if (!cs || cs.getLastRow() < 3) return toast_('Build the Count Sheet first.');
  const rows = cs.getRange(3, 1, cs.getLastRow() - 2, CS_HEAD.length).getValues();
  const totals = {}, extras = {};
  rows.forEach(r => {
    const id = String(r[2]);
    if (r[5] !== '' && !isNaN(r[5])) totals[id] = Number(r[5]);
    const cond = r[6] === '—' ? '' : r[6];
    const rep = r[7] === '—' ? '' : r[7];
    if (cond || rep !== '') extras[id] = { cond, replace: rep };
  });
  setOnHand_(totals, extras);
  toast_(`Updated ${Object.keys(totals).length} count(s). Blank counts were skipped.`);
}

// ───────────────────────────── Labels ─────────────────────────────
// 3 labels across (≈ Avery 5160). Barcode = Code 39 of the Item ID.

function buildLabels() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventory');
  if (inv.getLastRow() < 2) return toast_('Inventory is empty.');
  let items = inv.getRange(2, 1, inv.getLastRow() - 1, INV_HEAD.length).getValues().filter(v => v[INV.id - 1]);
  const checked = items.filter(v => v[INV.label - 1] === true);
  if (checked.length) items = checked;
  const lb = ensureSheet_(ss, 'Labels', []);
  lb.clear();
  const ACROSS = 3;
  for (let i = 0; i < items.length; i++) {
    const r = Math.floor(i / ACROSS) * 3 + 1;
    const c = (i % ACROSS) + 1;
    const v = items[i];
    const id = String(v[INV.id - 1]).toUpperCase();
    lb.getRange(r, c).setValue(String(v[INV.desc - 1]).slice(0, 38)).setFontSize(8);
    lb.getRange(r + 1, c).setValue('*' + id + '*').setFontFamily(BARCODE_FONT).setFontSize(30);
    lb.getRange(r + 2, c).setValue(v[INV.mode - 1] === 'Enter qty' ? '(scan, then enter qty)' : '').setFontSize(7);
  }
  lb.getRange(1, 1, lb.getMaxRows(), ACROSS).setHorizontalAlignment('center').setVerticalAlignment('middle');
  for (let c = 1; c <= ACROSS; c++) lb.setColumnWidth(c, 250);
  const lastRow = Math.ceil(items.length / ACROSS) * 3;
  for (let r = 1; r <= lastRow; r += 3) { lb.setRowHeight(r, 18); lb.setRowHeight(r + 1, 44); lb.setRowHeight(r + 2, 14); }
  ss.setActiveSheet(lb);
  toast_(`${items.length} label(s). Print on label stock or cardstock; tick "Label?" in Inventory to print only some.`);
}

// ───────────────────────────── Planning & ordering ─────────────────────────────

function syncPlanning() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventory');
  const pl = ensureSheet_(ss, 'Planning', PLAN_HEAD);
  const have = new Set(colValues_(pl, 1).map(String));
  const add = inv.getRange(2, 1, Math.max(inv.getLastRow() - 1, 1), INV_HEAD.length).getValues()
    .filter(v => v[INV.id - 1] && !have.has(String(v[INV.id - 1])))
    .map(v => [v[INV.id - 1], v[INV.desc - 1], v[INV.unit - 1], '', '', '', '', '']);
  if (!add.length) return toast_('Planning already lists every Inventory item.');
  const start = pl.getLastRow() + 1;
  pl.getRange(start, 1, add.length, 1).setNumberFormat('@');
  pl.getRange(start, 1, add.length, PLAN_HEAD.length).setValues(add);
  // Total Need = Per Student × Students + Extra (rounded up)
  pl.getRange(start, 7, add.length, 1)
    .setFormulaR1C1('=IF(AND(RC4="",RC6=""),"",ROUNDUP(N(RC4)*N(RC5)+N(RC6),0))');
  pl.getRange(1, 4, 1, 3).setNotes([[
    'How many units each student uses in a year (e.g. 0.25 = one per 4 students)',
    'Number of students who will use it',
    'Flat amount on top (teacher use, spares, or the whole need if not per-student)']]);
  pl.setFrozenRows(1); pl.setColumnWidth(2, 320);
  toast_(`Added ${add.length} item(s) to Planning.`);
}

function buildOrderList() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventory');
  const pl = ss.getSheetByName('Planning');
  if (!pl || pl.getLastRow() < 2) return toast_('Fill in Planning first.');
  const invById = {};
  inv.getRange(2, 1, Math.max(inv.getLastRow() - 1, 1), INV_HEAD.length).getValues()
    .forEach(v => invById[String(v[INV.id - 1])] = v);

  const out = [];
  pl.getRange(2, 1, pl.getLastRow() - 1, PLAN_HEAD.length).getValues().forEach(p => {
    const id = String(p[0]);
    const need = Number(p[6]) || 0;
    const v = invById[id];
    if (!id || !v) return;
    const onHand = Number(v[INV.onHand - 1]) || 0;
    const isDurable = v[INV.type - 1] === 'Durable';
    let bad = isDurable ? Number(v[INV.replace - 1]) || 0 : 0;
    if (isDurable && v[INV.cond - 1] === 'Replace' && !bad) bad = onHand;
    const usable = Math.max(0, onHand - bad);
    const qty = Math.max(0, Math.ceil(need - usable));
    if (!qty) return;
    const cost = v[INV.cost - 1];
    const reason = bad ? `Shortfall (incl. ${bad} worn out)` : 'Shortfall';
    out.push([id, v[INV.desc - 1], v[INV.unit - 1], need, usable, qty, cost,
              cost === '' ? '' : qty * cost, v[INV.source - 1], reason]);
  });

  const ol = ensureSheet_(ss, 'Order List', ORDER_HEAD);
  if (ol.getLastRow() > 1) ol.getRange(2, 1, ol.getLastRow() - 1, ol.getMaxColumns()).clear();
  if (!out.length) return toast_('Nothing to order — usable stock covers every planned need.');
  out.sort((a, b) => String(a[8]).localeCompare(String(b[8])) || String(a[0]).localeCompare(String(b[0])));
  ol.getRange(2, 1, out.length, 1).setNumberFormat('@');
  ol.getRange(2, 1, out.length, ORDER_HEAD.length).setValues(out);
  ol.getRange(2, 7, out.length, 2).setNumberFormat('$#,##0.00');
  const tot = out.length + 2;
  ol.getRange(tot, 7).setValue('TOTAL').setFontWeight('bold');
  ol.getRange(tot, 8).setFormula(`=SUM(H2:H${tot - 1})`).setNumberFormat('$#,##0.00').setFontWeight('bold');
  ol.setFrozenRows(1); ol.setColumnWidth(2, 320);
  ss.setActiveSheet(ol);
  toast_(`Order List: ${out.length} line(s). District items use the Stock # for FMS.`);
}

// ───────────────────────────── Helpers ─────────────────────────────

/** Picture formula reading a Drive file ID from a column `offset` to the right. */
function picFormula_(offset) {
  const ref = `RC[${offset}]`;
  return `=IF(${ref}="","",IMAGE("https://drive.google.com/thumbnail?id="&${ref}&"&sz=w200"))`;
}

function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (head.length && sh.getRange(1, 1).getValue() === '') {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#d9ead3');
    sh.setFrozenRows(1);
  }
  return sh;
}

function colValues_(sh, col) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(x => x !== '');
}

function list_(arr) {
  return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(false).build();
}

function toast_(msg) { SpreadsheetApp.getActive().toast(msg, 'Inventory', 8); }
