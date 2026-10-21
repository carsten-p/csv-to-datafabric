#!/usr/bin/env node
// Creates a UiPath Data Fabric entity from an arbitrary CSV file and imports its rows.
// Cross-platform (macOS, Linux, Windows), single file, no dependencies: needs only Node.js 18+,
// which the uip CLI already requires. Same behaviour as Import-CsvToDataFabric.ps1.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline/promises';

const HELP = `
Usage: node import-csv-to-datafabric.mjs <CSV_PATH> [options]

Creates a Data Fabric entity from the CSV's columns and imports its rows with the uip CLI.
Requires: uip on PATH, logged in (uip login), Data Fabric tool
(uip tools install @uipath/data-fabric-tool@latest).

Steps: parse CSV -> fix headers into valid field names (original header kept as display name)
-> infer types from all rows -> show plan, confirm -> uip df entities create -> read schema back
-> write cleaned CSV -> uip df records import (in chunks).

Options:
  --entity-name <NAME>          Entity name (default: from file name; letters, digits, underscores)
  --entity-display-name <TEXT>  Entity display name (default: file name)
  --description <TEXT>          Entity description
  --folder-key <GUID>           Orchestrator folder key for a folder-scoped entity (default: tenant)
                                Find it: uip or folders get "<FOLDER_PATH>" --output json
  --delimiter <CHAR>            , ; | or tab (default: auto-detect)
  --encoding <NAME>             Encoding when the file has no BOM: utf8 (default), 1252, latin1
  --decimal-comma               Numbers use a decimal comma (1.234,56 -> 1234.56)
  --date-formats <LIST>         Comma-separated date formats using yyyy, MM, dd
                                (default: yyyy-MM-dd,dd.MM.yyyy); values become yyyy-MM-dd
  --column-type <COL=TYPE>      Override a column type, repeatable. Key: original header or field name.
                                TYPE: STRING, MULTILINE_TEXT, MULTILINE_MAX, BOOLEAN, DATE,
                                DATETIME_WITH_TZ, DECIMAL or DECIMAL:<precision>
  --all-strings                 Skip type inference; every column becomes text
  --plain-display-names         Use the sanitized field name as display name
  --use-existing                Import into an existing entity of that name; error if it does not exist
  --schema-only                 Create the entity, import no rows
  --chunk-size <N>              Rows per records import call (default 5000, 0 = one call)
  --force, -y                   Do not ask for confirmation (required when not run in a terminal)
  --dry-run                     Show the plan, write entity JSON + cleaned CSV, change nothing
  --keep-temp                   Keep generated files
  --help, -h                    Show this help

Examples:
  node import-csv-to-datafabric.mjs ./customers.csv --dry-run
  node import-csv-to-datafabric.mjs ./Kunden.csv --encoding 1252 --decimal-comma --column-type PLZ=STRING
  node import-csv-to-datafabric.mjs ./orders.csv --entity-name Orders_2026 --folder-key <FOLDER_KEY> --force

Progress goes to stderr; the result summary is printed to stdout as JSON.
`;

// --- Arguments -------------------------------------------------------------------------------

const opts = {
  csv: null, entityName: null, entityDisplayName: null, description: '', folderKey: null,
  delimiter: null, encoding: 'utf8', decimalComma: false, dateFormats: ['yyyy-MM-dd', 'dd.MM.yyyy'],
  columnTypes: {}, allStrings: false, plainDisplayNames: false, useExisting: false, schemaOnly: false,
  chunkSize: 5000, force: false, dryRun: false, keepTemp: false,
};
const FLAGS = {
  '--decimal-comma': 'decimalComma', '--all-strings': 'allStrings', '--plain-display-names': 'plainDisplayNames',
  '--use-existing': 'useExisting', '--schema-only': 'schemaOnly', '--force': 'force', '-y': 'force',
  '--dry-run': 'dryRun', '--keep-temp': 'keepTemp',
};
const VALUED = {
  '--csv': 'csv', '--entity-name': 'entityName', '--entity-display-name': 'entityDisplayName',
  '--description': 'description', '--folder-key': 'folderKey', '--delimiter': 'delimiter',
  '--encoding': 'encoding', '--date-formats': 'dateFormats', '--column-type': 'columnType', '--chunk-size': 'chunkSize',
};

function parseArgs(argv) {
  for (let i = 0; i < argv.length; i++) {
    let a = argv[i];
    let inline = null;
    if (a.startsWith('--') && a.includes('=')) { inline = a.slice(a.indexOf('=') + 1); a = a.slice(0, a.indexOf('=')); }
    if (a === '--help' || a === '-h') { process.stdout.write(HELP); process.exit(0); }
    if (FLAGS[a]) { opts[FLAGS[a]] = true; continue; }
    if (VALUED[a]) {
      const v = inline !== null ? inline : argv[++i];
      if (v === undefined) fail(`Missing value for ${a}`);
      const key = VALUED[a];
      if (key === 'columnType') {
        const eq = v.lastIndexOf('=');
        if (eq < 1) fail(`--column-type expects COLUMN=TYPE, got '${v}'`);
        opts.columnTypes[v.slice(0, eq)] = v.slice(eq + 1);
      } else if (key === 'dateFormats') opts.dateFormats = v.split(',').map(s => s.trim()).filter(Boolean);
      else if (key === 'chunkSize') opts.chunkSize = parseInt(v, 10);
      else opts[key] = v;
      continue;
    }
    if (a.startsWith('-')) fail(`Unknown option ${a}. Use --help.`);
    if (opts.csv) fail(`Unexpected argument '${a}'`);
    opts.csv = a;
  }
  if (!opts.csv) { process.stdout.write(HELP); process.exit(1); }
  if (opts.delimiter && /^(tab|\\t)$/i.test(opts.delimiter)) opts.delimiter = '\t';
}

function fail(msg) { throw new Error(msg); }
const log = (...m) => console.error(...m);
const warn = m => console.error(`WARNING: ${m}`);

// --- Name rules -------------------------------------------------------------------------------

const SYSTEM_FIELDS = ['Id', 'CreatedBy', 'CreateTime', 'UpdatedBy', 'UpdateTime', 'RecordOwner'];
const SYSTEM_DISPLAY = ['Id', 'Created By', 'Create Time', 'Updated By', 'Update Time', 'Record Owner', ...SYSTEM_FIELDS];
const lc = s => s.toLowerCase();
const isSystemField = s => SYSTEM_FIELDS.some(f => lc(f) === lc(s));
const isSystemDisplay = s => SYSTEM_DISPLAY.some(f => lc(f) === lc(s));

// C# + VB.NET keywords and SQL words the service rejects. The server stays the final judge;
// renaming a harmless name only changes the internal field name, the display name keeps the header.
const KEYWORDS = new Set((
  'abstract as base bool break byte case catch char checked class const continue decimal default delegate do ' +
  'double else enum event explicit extern false finally fixed float for foreach goto if implicit in int interface ' +
  'internal is lock long namespace new null object operator out override params private protected public readonly ' +
  'ref return sbyte sealed short sizeof stackalloc static string struct switch this throw true try typeof uint ulong ' +
  'unchecked unsafe ushort using virtual void volatile while addhandler addressof alias and andalso boolean byref ' +
  'byval call cbool cbyte cchar cdate cdbl cdec cint clng cobj csbyte cshort csng cstr ctype cuint culng cushort date ' +
  'declare dim directcast each elseif end endif erase error exit friend function get gettype getxmlnamespace global ' +
  'gosub handles implements imports inherits integer isnot let lib like loop me mod module mustinherit mustoverride ' +
  'mybase myclass narrowing next not nothing notinheritable notoverridable of on option optional or orelse overloads ' +
  'overridable overrides paramarray partial property raiseevent redim rem removehandler resume select set shadows ' +
  'shared single step stop sub synclock then to trycast variant wend when widening with withevents writeonly xor ' +
  'order group table from where by index view').split(' '));

const ASCII_MAP = [
  ['ä', 'ae'], ['ö', 'oe'], ['ü', 'ue'], ['Ä', 'Ae'], ['Ö', 'Oe'], ['Ü', 'Ue'],
  ['ß', 'ss'], ['€', 'Eur'], ['%', 'Pct'], ['#', 'No'],
];
function toAscii(s) {
  for (const [from, to] of ASCII_MAP) s = s.split(from).join(to);
  return s.normalize('NFD').replace(/\p{M}/gu, '');
}

// Header -> PascalCase letters+digits; allowUnderscore for entity names.
function safeName(raw, allowUnderscore, fallback) {
  let s = toAscii(raw);
  if (allowUnderscore) {
    s = s.replace(/[^A-Za-z0-9_]+/g, '_').replace(/^_+|_+$/g, '');
  } else {
    // ALL-CAPS tokens (SAP-style COMPANY_CODE) become CompanyCode; the server would otherwise store Companycode.
    s = s.split(/[^A-Za-z0-9]+/).filter(Boolean)
      .map(p => (/^[A-Z0-9]+$/.test(p) && p.length > 1 ? p.toLowerCase() : p))
      .map(p => p[0].toUpperCase() + p.slice(1)).join('');
  }
  if (!s) s = fallback;
  if (!/^[A-Za-z]/.test(s)) s = 'F' + s;
  if (s.length < 3) s += 'Value';
  if (s.length > 100) s = s.slice(0, 100);
  return s;
}

function uniqueName(name, taken, maxLen, separator = '') {
  let candidate = name;
  for (let i = 2; taken.has(lc(candidate)); i++) {
    const suffix = separator + i;
    candidate = name.slice(0, Math.min(name.length, maxLen - suffix.length)) + suffix;
  }
  taken.add(lc(candidate));
  return candidate;
}

// --- CSV reading ------------------------------------------------------------------------------

function decodeFile(file, fallback) {
  const buf = fs.readFileSync(file);
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return { text: buf.subarray(3).toString('utf8'), name: 'utf-8' };
  if (buf[0] === 0xff && buf[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(buf.subarray(2)), name: 'utf-16le' };
  if (buf[0] === 0xfe && buf[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(buf.subarray(2)), name: 'utf-16be' };
  const label = { utf8: 'utf-8', 'utf-8': 'utf-8', 1252: 'windows-1252', 'windows-1252': 'windows-1252', latin1: 'latin1' }[lc(String(fallback))] || fallback;
  return { text: new TextDecoder(label).decode(buf), name: label };
}

function detectDelimiter(text) {
  const firstLine = text.split(/\r?\n/, 1)[0];
  let best = ',', bestCount = -1;
  for (const d of [',', ';', '\t', '|']) {
    let inQuote = false, count = 0;
    for (const ch of firstLine) {
      if (ch === '"') inQuote = !inQuote;
      else if (!inQuote && ch === d) count++;
    }
    if (count > bestCount) { best = d; bestCount = count; }
  }
  return best;
}

// RFC 4180: quoted fields, "" escapes, delimiters and line breaks inside quotes; blank lines skipped.
function parseCsv(text, delim) {
  const rows = [];
  let row = [], field = '', inQuote = false, line = 1, quoteStart = 0;
  const endRow = () => { row.push(field); field = ''; if (!(row.length === 1 && row[0] === '')) rows.push(row); row = []; };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuote) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuote = false;
      } else { if (ch === '\n') line++; field += ch; }
    } else if (ch === '"' && field === '') { inQuote = true; quoteStart = line; }
    else if (ch === delim) { row.push(field); field = ''; }
    else if (ch === '\r' || ch === '\n') { if (ch === '\r' && text[i + 1] === '\n') i++; line++; endRow(); }
    else field += ch;
  }
  if (inQuote) fail(`Malformed CSV: quoted value starting at line ${quoteStart} is never closed.`);
  if (field !== '' || row.length) endRow();
  return rows;
}

// --- Type inference ---------------------------------------------------------------------------

const MAX_SAFE = 9007199254740991;
const DATETIME_RE = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d{1,7})?)?(Z|[+-]\d{2}:?\d{2})?$/;
const notEmpty = v => v != null && v.trim() !== '';

function toInvariantNumber(v) {
  let t = v.trim();
  if (opts.decimalComma) {
    if (!/^[+-]?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?$/.test(t)) return null;
    t = t.replace(/\./g, '').replace(',', '.');
  } else if (!/^[+-]?\d+(\.\d+)?$/.test(t)) return null;
  // Leading zeros (ZIP codes, article numbers) must stay text.
  if (/^[+-]?0\d/.test(t)) return null;
  if (Math.abs(Number(t)) > MAX_SAFE) return null;
  return t;
}

let DATE_PARSERS = [];
function compileDateFormats() {
  DATE_PARSERS = opts.dateFormats.map(f => {
    const order = [];
    const src = f.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')
      .replace(/yyyy|MM|dd/g, m => { order.push(m); return m === 'yyyy' ? '(\\d{4})' : '(\\d{2})'; });
    if (!['yyyy', 'MM', 'dd'].every(p => order.includes(p))) fail(`Date format '${f}' must contain yyyy, MM and dd.`);
    return { re: new RegExp(`^${src}$`), order };
  });
}

function toIsoDate(v) {
  const t = v.trim();
  for (const { re, order } of DATE_PARSERS) {
    const m = re.exec(t);
    if (!m) continue;
    const part = {};
    order.forEach((p, i) => { part[p] = parseInt(m[i + 1], 10); });
    const d = new Date(Date.UTC(part.yyyy, part.MM - 1, part.dd));
    if (d.getUTCFullYear() === part.yyyy && d.getUTCMonth() === part.MM - 1 && d.getUTCDate() === part.dd) {
      return d.toISOString().slice(0, 10);
    }
  }
  return null;
}

function toIsoDateTime(v) {
  let t = v.trim();
  if (!DATETIME_RE.test(t)) return null;
  t = t.replace(' ', 'T').replace(/(\.\d{3})\d+/, '$1').replace(/([+-]\d{2})(\d{2})$/, '$1:$2');
  const d = new Date(t); // no offset -> local time, like the PowerShell version
  return isNaN(d.getTime()) ? null : d.toISOString();
}

function stringType(maxLen) {
  if (maxLen <= 4000) return { type: 'STRING', lengthLimit: Math.min(4000, Math.max(255, Math.ceil(maxLen * 1.5))) };
  if (maxLen <= 10000) return { type: 'MULTILINE_TEXT', lengthLimit: 10000 };
  return { type: 'MULTILINE_MAX' };
}

function inferType(values, maxLen) {
  const filled = values.filter(notEmpty);
  if (opts.allStrings || filled.length === 0) return stringType(maxLen);
  if (filled.every(v => /^(true|false)$/i.test(v.trim()))) return { type: 'BOOLEAN' };
  let precision = 0;
  const isNumber = filled.every(v => {
    const n = toInvariantNumber(v);
    if (n === null) return false;
    const dot = n.indexOf('.');
    if (dot >= 0) precision = Math.max(precision, n.length - dot - 1);
    return true;
  });
  if (isNumber && precision <= 10) return { type: 'DECIMAL', decimalPrecision: precision };
  if (filled.every(v => toIsoDate(v) !== null)) return { type: 'DATE' };
  if (filled.every(v => toIsoDateTime(v) !== null)) return { type: 'DATETIME_WITH_TZ' };
  return stringType(maxLen);
}

function typeOverride(spec, maxLen) {
  const s = spec.trim().toUpperCase();
  const dec = /^DECIMAL(:(\d+))?$/.exec(s);
  if (dec) return { type: 'DECIMAL', decimalPrecision: dec[2] ? parseInt(dec[2], 10) : 0 };
  if (s === 'STRING') return { type: 'STRING', lengthLimit: Math.min(4000, Math.max(255, Math.ceil(maxLen * 1.5))) };
  if (s === 'MULTILINE_TEXT') return { type: 'MULTILINE_TEXT', lengthLimit: 10000 };
  if (['MULTILINE_MAX', 'BOOLEAN', 'DATE', 'DATETIME_WITH_TZ'].includes(s)) return { type: s };
  fail(`Unsupported type override '${spec}'. Use STRING, MULTILINE_TEXT, MULTILINE_MAX, BOOLEAN, DATE, DATETIME_WITH_TZ, DECIMAL[:n].`);
}

function importValue(v, type) {
  if (!notEmpty(v)) return '';
  if (type === 'BOOLEAN') return v.trim().toLowerCase();
  if (type === 'DECIMAL') return toInvariantNumber(v) ?? v;
  if (type === 'DATE') return toIsoDate(v) ?? v;
  if (type === 'DATETIME_WITH_TZ') return toIsoDateTime(v) ?? v;
  return v;
}

// Typed JSON value for records insert; undefined = leave the field empty.
function insertValue(v, type) {
  const s = importValue(v, type);
  if (s === '') return undefined;
  if (type === 'BOOLEAN' && /^(true|false)$/.test(s)) return s === 'true';
  if (type === 'DECIMAL' && toInvariantNumber(s) !== null) return Number(s);
  return s;
}

const csvLine = fields => fields.map(f => `"${String(f).replace(/"/g, '""')}"`).join(',');

// --- uip wrapper ------------------------------------------------------------------------------

const winQuote = a => (/^[A-Za-z0-9_\-.:\\/=@]+$/.test(a) ? a : `"${a.replace(/"/g, '""')}"`);

function uip(args) {
  const all = [...args, '--output', 'json'];
  const spawnOpts = { encoding: 'utf8', maxBuffer: 512 * 1024 * 1024 };
  // Windows: uip is a .cmd shim, which Node only runs through a shell.
  const r = process.platform === 'win32'
    ? spawnSync(['uip', ...all].map(winQuote).join(' '), { ...spawnOpts, shell: true })
    : spawnSync('uip', all, spawnOpts);
  if (r.error) {
    if (r.error.code === 'ENOENT') fail('uip CLI not found on PATH. Install the UiPath CLI and run uip login.');
    throw r.error;
  }
  // uip writes notices ("Checking today's update...") to stderr; the JSON envelope is on stdout.
  const out = r.stdout || '';
  const start = out.indexOf('{'), end = out.lastIndexOf('}');
  if (start < 0 || end < start) fail(`uip ${args.join(' ')} returned no JSON (exit ${r.status}):\n${out}${r.stderr || ''}`);
  const json = JSON.parse(out.slice(start, end + 1));
  if (json.Result !== 'Success') fail(`uip ${args.slice(0, 3).join(' ')} failed: ${out.slice(start, end + 1)}`);
  return json;
}
const withFolder = args => (opts.folderKey ? [...args, '--folder-key', opts.folderKey] : args);

// --- Main -------------------------------------------------------------------------------------

function printPlan(columns) {
  const rows = columns.map(c => {
    let t = c.type;
    if ('decimalPrecision' in c.typeInfo) t += `:${c.typeInfo.decimalPrecision}`;
    else if ('lengthLimit' in c.typeInfo) t += `(${c.typeInfo.lengthLimit})`;
    return [c.header, c.name, c.displayName, t, c.source, c.sample];
  });
  const head = ['Header', 'Name', 'DisplayName', 'Type', 'Source', 'Sample'];
  const width = head.map((h, i) => Math.max(h.length, ...rows.map(r => r[i].length)));
  const fmt = r => r.map((v, i) => v.padEnd(width[i])).join('  ');
  log(fmt(head));
  log(fmt(width.map(w => '-'.repeat(w))));
  rows.forEach(r => log(fmt(r)));
  log('');
}

async function confirm(question) {
  if (!process.stdin.isTTY) fail('Not running in a terminal - pass --force to create the entity without confirmation.');
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  try { return /^(y|yes|j|ja)$/i.test((await rl.question(question)).trim()); } finally { rl.close(); }
}

async function main() {
  parseArgs(process.argv.slice(2));
  compileDateFormats();

  const csvPath = path.resolve(opts.csv);
  if (!fs.existsSync(csvPath)) fail(`CSV '${csvPath}' not found.`);
  const { text, name: encName } = decodeFile(csvPath, opts.encoding);
  const delimiter = opts.delimiter || detectDelimiter(text);
  const rows = parseCsv(text, delimiter);
  if (rows.length < 1) fail(`CSV '${csvPath}' is empty.`);
  const headers = rows[0];
  const dataRows = rows.slice(1);
  log(`Read ${dataRows.length} data rows, ${headers.length} columns (delimiter '${delimiter === '\t' ? 'TAB' : delimiter}', encoding ${encName}).`);

  const badRows = dataRows.map((r, i) => (r.length !== headers.length ? i + 2 : 0)).filter(Boolean);
  if (badRows.length) warn(`${badRows.length} row(s) have a different column count than the header (rows: ${badRows.slice(0, 10).join(', ')}). Missing cells are imported empty, extra cells are dropped.`);
  if (text.includes('�')) warn('The file contains undecodable characters. If it is an Excel export, retry with --encoding 1252.');

  const fileBase = path.basename(csvPath, path.extname(csvPath));
  const entityName = opts.entityName || safeName(fileBase, true, 'CsvImport');
  if (!/^[A-Za-z][A-Za-z0-9_]{2,99}$/.test(entityName)) fail(`Entity name '${entityName}' is invalid: start with a letter, letters/digits/underscores, 3-100 chars.`);
  const entityDisplayName = opts.entityDisplayName || fileBase;

  // Column plan
  const takenNames = new Set([...SYSTEM_FIELDS, entityName].map(lc));
  const takenDisplay = new Set(SYSTEM_DISPLAY.map(lc));
  const columns = headers.map((rawHeader, c) => {
    const header = rawHeader.replace(/^﻿/, '');
    const cleanHeader = header.replace(/\s+/g, ' ').trim();

    let name = safeName(cleanHeader, false, `Column${c + 1}`);
    if (isSystemField(name)) name = 'Source' + name;
    else if (KEYWORDS.has(lc(name)) || lc(name) === lc(entityName)) name += 'Value';
    name = uniqueName(name, takenNames, 100);

    let display = cleanHeader;
    if (opts.plainDisplayNames || display === '') display = name;
    if (isSystemDisplay(display)) display = 'Source ' + display;
    if (display.length > 120) display = display.slice(0, 120);
    display = uniqueName(display, takenDisplay, 128, ' ');

    const values = dataRows.map(r => (c < r.length ? r[c] : ''));
    const maxLen = values.reduce((m, v) => Math.max(m, v.length), 0);
    const overrideKey = [header, cleanHeader, name].find(k => k && k in opts.columnTypes);
    const typeInfo = overrideKey ? typeOverride(opts.columnTypes[overrideKey], maxLen) : inferType(values, maxLen);
    let sample = values.find(notEmpty) || '';
    sample = sample.replace(/\s+/g, ' ');
    if (sample.length > 30) sample = sample.slice(0, 27) + '...';

    return { index: c, header, name, displayName: display, type: typeInfo.type, typeInfo,
      source: overrideKey ? 'override' : 'inferred', sample, importHeader: null, skip: false };
  });
  const unknownOverrides = Object.keys(opts.columnTypes).filter(k => !columns.some(c => [c.header, c.name].includes(k) || c.header.replace(/\s+/g, ' ').trim() === k));
  if (unknownOverrides.length) warn(`--column-type for unknown column(s): ${unknownOverrides.join(', ')}`);

  log('');
  log(`Entity: ${entityName}  (display '${entityDisplayName}', scope: ${opts.folderKey ? `folder ${opts.folderKey}` : 'tenant'})`);
  log('');
  printPlan(columns);

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'df-import-'));
  const fields = columns.map(c => {
    const f = { name: c.name, type: c.type, displayName: c.displayName };
    for (const k of ['lengthLimit', 'decimalPrecision']) if (k in c.typeInfo) f[k] = c.typeInfo[k];
    return f;
  });
  const bodyPath = path.join(tempDir, 'entity.json');
  fs.writeFileSync(bodyPath, JSON.stringify({ displayName: entityDisplayName, description: opts.description, fields }, null, 2));

  // records import ends a row at every line break, even inside quotes. Rows with a multi-line value
  // therefore go through records insert (JSON) instead; all other rows use the bulk CSV import.
  const cell = (row, c) => (c.index < row.length ? row[c.index] : '');
  const isMultiline = row => columns.some(c => !c.skip && /[\r\n]/.test(cell(row, c)));

  const writeImportFiles = rowsToWrite => {
    const active = columns.filter(c => !c.skip);
    const header = csvLine(active.map(c => c.importHeader));
    const size = opts.chunkSize > 0 ? opts.chunkSize : Math.max(1, rowsToWrite.length);
    const files = [];
    for (let start = 0; start < rowsToWrite.length; start += size) {
      const lines = [header];
      for (const row of rowsToWrite.slice(start, start + size)) {
        lines.push(csvLine(active.map(c => importValue(cell(row, c), c.type))));
      }
      const file = path.join(tempDir, `import-${String(files.length + 1).padStart(4, '0')}.csv`);
      fs.writeFileSync(file, lines.join('\r\n') + '\r\n', 'utf8');
      files.push(file);
    }
    return files;
  };
  const cleanup = keep => {
    if (keep) log(`Generated files kept in ${tempDir}`);
    else fs.rmSync(tempDir, { recursive: true, force: true });
  };

  const multiRows = dataRows.filter(isMultiline);
  const bulkRows = dataRows.filter(r => !isMultiline(r));
  if (multiRows.length) log(`${multiRows.length} row(s) contain line breaks inside a value; they are inserted one batch at a time with records insert.`);

  if (opts.dryRun) {
    columns.forEach(c => { c.importHeader = c.displayName; c.fieldName = c.name; });
    const files = writeImportFiles(bulkRows);
    log('Dry run - nothing changed in the tenant.');
    log(`Entity definition: ${bodyPath}`);
    log(`Cleaned CSV:       ${files.join(', ') || '(none)'}`);
    if (multiRows.length) log(`Multi-line rows:   ${multiRows.length} (sent via records insert)`);
    return;
  }

  // Pre-flight: login + existing entity
  const login = uip(['login', 'status']);
  log(`Tenant: ${login.Data.Organization}/${login.Data.Tenant} as ${login.Data.UserName}`);

  const existing = (uip(withFolder(['df', 'entities', 'list', '--native-only'])).Data || []).filter(e => lc(e.Name) === lc(entityName));
  let entityId, created = false;
  if (existing.length) {
    if (!opts.useExisting) fail(`Entity '${entityName}' already exists (Id ${existing[0].Id}). Pass --use-existing to import into it, or --entity-name for a new one.`);
    entityId = existing[0].Id;
    log(`Using existing entity ${entityName} (${entityId}).`);
  } else {
    if (opts.useExisting) {
      fail(`Entity '${entityName}' does not exist ${opts.folderKey ? `in folder ${opts.folderKey}` : 'at tenant level'}. --use-existing never creates an entity; check --entity-name / --folder-key, or drop --use-existing to create it.`);
    }
    if (!opts.force) {
      log('Field types cannot be changed after the entity is created.');
      const what = opts.schemaOnly ? '(schema only, no rows)' : `and import ${dataRows.length} rows`;
      if (!(await confirm(`Create entity '${entityName}' with ${columns.length} fields ${what}? [y/N] `))) {
        log(`Aborted. Entity definition left at ${bodyPath}`);
        return;
      }
    }
    entityId = uip(withFolder(['df', 'entities', 'create', entityName, '--file', bodyPath])).Data.Id;
    created = true;
    log(`Created entity ${entityName} (${entityId}).`);
  }

  const summary = { EntityName: entityName, EntityId: entityId, Created: created, FolderKey: opts.folderKey, RowsInCsv: dataRows.length };
  if (opts.schemaOnly) {
    if (!created) log('Schema only: entity already existed, nothing changed.');
    cleanup(opts.keepTemp);
    console.log(JSON.stringify({ ...summary, InsertedRecords: 0 }, null, 2));
    return;
  }

  // Read the schema back: the import header must be the server's DisplayName for each field.
  const schema = uip(withFolder(['df', 'entities', 'get', entityId])).Data;
  const serverFields = (schema.Fields || []).filter(f => !isSystemField(f.Name));
  const COMPLEX = ['CHOICE_SET_SINGLE', 'CHOICE_SET_MULTIPLE', 'RELATIONSHIP', 'FILE', 'AUTO_NUMBER'];
  for (const c of columns) {
    const sf = serverFields.find(f => lc(f.Name) === lc(c.name))
      || serverFields.find(f => lc(f.DisplayName) === lc(c.displayName) || lc(f.DisplayName) === lc(c.header));
    if (!sf) { warn(`Column '${c.header}' has no matching field in the entity - skipped.`); c.skip = true; continue; }
    const t = sf.FieldDataType && sf.FieldDataType.Name;
    if (COMPLEX.includes(t)) { warn(`Field '${sf.Name}' is ${t}; records import ignores this type - skipped.`); c.skip = true; continue; }
    c.importHeader = sf.DisplayName;
    c.fieldName = sf.Name;
    c.type = t;
  }
  const mapped = columns.filter(c => !c.skip).map(c => c.importHeader);
  for (const sf of serverFields) {
    if (sf.IsRequired && !mapped.includes(sf.DisplayName) && !sf.DefaultValue) warn(`Required field '${sf.Name}' is not in the CSV - every row will fail.`);
  }
  if (!mapped.length) fail(`No CSV column maps to a field of entity ${entityName}.`);

  const files = writeImportFiles(bulkRows);
  let inserted = 0, total = 0;
  const errorLinks = [], failedChunks = [];
  const insertFailures = [];
  for (const [i, file] of files.entries()) {
    log(`Importing chunk ${i + 1}/${files.length}...`);
    try {
      const d = uip(withFolder(['df', 'records', 'import', entityId, '--file', file])).Data;
      inserted += Number(d.InsertedRecords || 0);
      total += Number(d.TotalRecords || 0);
      if (d.ErrorFileLink) errorLinks.push(d.ErrorFileLink);
    } catch (e) {
      warn(`Chunk ${i + 1} failed: ${e.message}`);
      failedChunks.push(file);
    }
  }

  const INSERT_BATCH = 100;
  const active = columns.filter(c => !c.skip);
  for (let start = 0; start < multiRows.length; start += INSERT_BATCH) {
    const records = multiRows.slice(start, start + INSERT_BATCH).map(row => {
      const rec = {};
      for (const c of active) { const v = insertValue(cell(row, c), c.type); if (v !== undefined) rec[c.fieldName] = v; }
      return rec;
    });
    const file = path.join(tempDir, `insert-${String(start / INSERT_BATCH + 1).padStart(4, '0')}.json`);
    fs.writeFileSync(file, JSON.stringify(records), 'utf8');
    log(`Inserting multi-line rows ${start + 1}-${start + records.length}...`);
    total += records.length;
    try {
      const d = uip(withFolder(['df', 'records', 'insert', entityId, '--file', file])).Data;
      if (records.length === 1) inserted += 1;
      else {
        inserted += Number(d.SuccessCount || 0);
        if (Number(d.FailureCount || 0) > 0) insertFailures.push(...(d.FailureRecords || []));
      }
    } catch (e) {
      warn(`Insert batch starting at multi-line row ${start + 1} failed: ${e.message}`);
      failedChunks.push(file);
    }
  }
  let insertFailureFile = null;
  if (insertFailures.length) {
    insertFailureFile = path.join(tempDir, 'insert-failures.json');
    fs.writeFileSync(insertFailureFile, JSON.stringify(insertFailures, null, 2), 'utf8');
    warn(`${insertFailures.length} multi-line row(s) were rejected; details in ${insertFailureFile}`);
  }

  const incomplete = failedChunks.length > 0 || errorLinks.length > 0 || insertFailures.length > 0 || inserted < dataRows.length;
  if (incomplete) {
    warn(`Not every row was imported. Download the ErrorFileLinks for row errors. Re-run a failed CSV chunk with: uip df records import ${entityId} --file <chunk>${opts.folderKey ? ` --folder-key ${opts.folderKey}` : ''} --output json (insert-*.json files: uip df records insert ... --file <file>)`);
  }
  cleanup(opts.keepTemp || incomplete);
  console.log(JSON.stringify({ ...summary, InsertedRecords: inserted, TotalRecords: total, ErrorFileLinks: errorLinks, InsertFailureFile: insertFailureFile, FailedChunks: failedChunks }, null, 2));
  if (incomplete) process.exitCode = 2;
}

main().catch(e => { console.error(`ERROR: ${e.message}`); process.exit(1); });
