/**
 * Validates every template under `templates/` and writes the catalog `dist/` that unpkg serves.
 *
 *   templates/<category>/<name>.xlsx|.sxl|.csv a template; its title is the file name, title-cased
 *   templates/<category>/<name>.txt            optional plain-text description, the gallery tooltip
 *   templates/<category>/<name>.webp|.png      optional hand-made thumbnail, copied through
 *   templates/<category>/<name>.dark.webp|.png optional hand-made dark thumbnail
 *   .thumbnails/<category>/<name>[.dark].webp   the rendered thumbnails (`npm run thumbnails`), used
 *                                               when there is no hand-made one
 *   templates/<category>/_category.json        optional: { title?, description?, order?, icon? }
 *
 * Output:
 *
 *   dist/contents.json                              root listing of every file and folder
 *   dist/workbook-templates/contents.json           the catalog
 *   dist/workbook-templates/<category>/<name>.xlsx|.sxl|.csv
 *   dist/workbook-templates/<category>/<name>.webp|.png
 *   dist/workbook-templates/<category>/<name>.dark.webp|.png
 *
 * A template with a problem is left out of the catalog, and a category with a problem is left out
 * with its templates; every problem is listed as a warning and the build succeeds. Only a catalog
 * with no templates at all fails it. `--strict` fails the build on any problem instead.
 */
import fs from "fs/promises";
import path from "path";
import zlib from "zlib";
import { WorkbookIO } from "@sheetxl/sdk";
import type { IWorkbook } from "@sheetxl/sdk";
import { IOPlugin } from "@sheetxl/io";

await WorkbookIO.install(IOPlugin);

const PACKAGE = "workbook-templates";
const SCHEMA = 1;
const MAX_BYTES = 2 * 1024 * 1024;
/** The template formats, by file extension, and the SDK format each is read as. */
const TEMPLATE_FORMATS: Record<string, string> = { ".xlsx": "xlsx", ".sxl": "sxl", ".csv": "csv" };
const TEMPLATE_EXTENSIONS = Object.keys(TEMPLATE_FORMATS);
const THUMBNAIL_EXTENSIONS = [".webp", ".png"];
const DESCRIPTION_EXTENSION = ".txt";
/** A description is a tooltip, not documentation. */
const MAX_DESCRIPTION_LENGTH = 300;
/** The file-name suffix of a dark thumbnail, before the extension. */
const DARK_SUFFIX = ".dark";

const STRICT = process.argv.includes("--strict");

const SRC = path.resolve("templates");
const RENDERED = path.resolve(".thumbnails");
const DIST = path.resolve("dist");
const OUT = path.join(DIST, PACKAGE);

/** Content gate: a template containing any of these is left out. */
const GATE = {
  scripts: "script modules",
  vba: "a VBA project",
  "external-links": "links to other workbooks",
  fetch: "functions that fetch from the network",
  "http-links": "hyperlinks that are not https",
  size: `a file over ${MAX_BYTES / 1024 / 1024} MB`,
} as const;
type GateItem = keyof typeof GATE;

/** Formula functions that reach the network (Excel and Google Sheets spellings). */
const FETCH_FUNCTIONS = [
  "WEBSERVICE", "FILTERXML", "RTD", "CALL", "REGISTER.ID", "STOCKHISTORY", "IMAGE",
  "IMPORTDATA", "IMPORTXML", "IMPORTHTML", "IMPORTFEED", "IMPORTRANGE", "GOOGLEFINANCE",
];
const FETCH_REGEX = new RegExp(`(?<![A-Za-z0-9_.])(?:_xlfn\\.|_xludf\\.)?(${FETCH_FUNCTIONS.map((f) => f.replace(".", "\\.")).join("|")})\\s*\\(`, "i");
const HTTP_LITERAL_REGEX = /["']\s*(?:http|ftp|file):/i;

interface CategoryFile {
  title?: string;
  description?: string;
  order?: number;
  icon?: string;
}

interface Category {
  id: string;
  title: string;
  description?: string;
  icon?: string;
  order?: number;
}

interface Entry {
  path: string;
  category: string;
  title: string;
  description?: string;
  thumbnail?: string;
  /** The thumbnail on the dark grid. */
  thumbnailDark?: string;
}

const problems: { where: string; message: string }[] = [];
function fail(where: string, message: string): void {
  problems.push({ where, message });
}

/** "yearly-calendar" becomes "Yearly Calendar". */
function toName(id: string): string {
  return id
    .split(/[-_\s]+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

const posix = (p: string) => p.split(path.sep).join("/");

async function exists(file: string): Promise<boolean> {
  return fs.access(file).then(() => true, () => false);
}

// ---------------------------------------------------------------------------------------------
// JSON inputs
// ---------------------------------------------------------------------------------------------

async function readJSON(file: string): Promise<unknown> {
  try {
    return JSON.parse(await fs.readFile(file, "utf-8"));
  } catch (err: any) {
    fail(posix(path.relative(process.cwd(), file)), `invalid JSON (${err.message})`);
    return undefined;
  }
}

function checkKeys(where: string, json: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(json)) {
    if (!allowed.includes(key)) fail(where, `unknown key "${key}" (allowed: ${allowed.join(", ")})`);
  }
}

function optString(where: string, json: Record<string, unknown>, key: string): string | undefined {
  const v = json[key];
  if (v === undefined) return undefined;
  if (typeof v !== "string" || !v.trim()) {
    fail(where, `"${key}" must be a non-empty string`);
    return undefined;
  }
  return v.trim();
}

async function readCategory(file: string): Promise<CategoryFile> {
  const where = posix(path.relative(process.cwd(), file));
  const json = await readJSON(file);
  if (json === undefined) return {};
  if (typeof json !== "object" || json === null || Array.isArray(json)) {
    fail(where, "must be a JSON object");
    return {};
  }
  const obj = json as Record<string, unknown>;
  checkKeys(where, obj, ["title", "description", "order", "icon"]);
  if (obj.order !== undefined && (typeof obj.order !== "number" || !Number.isFinite(obj.order))) {
    fail(where, `"order" must be a number`);
  }
  return {
    title: optString(where, obj, "title"),
    description: optString(where, obj, "description"),
    order: typeof obj.order === "number" ? obj.order : undefined,
    icon: optString(where, obj, "icon"),
  };
}

/** A description sidecar: plain text, whitespace collapsed to single spaces. */
async function readDescription(file: string): Promise<string | undefined> {
  const where = posix(path.relative(process.cwd(), file));
  const text = (await fs.readFile(file, "utf-8")).replace(/^﻿/, "").replace(/\s+/g, " ").trim();
  if (!text) {
    fail(where, "description is empty");
    return undefined;
  }
  if (/<\/?[a-z][^>]*>/i.test(text)) fail(where, "descriptions are plain text; the gallery shows markup as-is");
  if (text.length > MAX_DESCRIPTION_LENGTH) {
    fail(where, `description is ${text.length} characters; keep it under ${MAX_DESCRIPTION_LENGTH}`);
  }
  return text;
}

// ---------------------------------------------------------------------------------------------
// Package-level inspection: the .xlsx is a zip, and some parts are cheaper to see there
// ---------------------------------------------------------------------------------------------

interface ZipEntry {
  name: string;
  read(): Buffer;
}

/** Lists a zip's entries from its central directory. Enough for small OOXML packages; no zip64. */
function readZip(buf: Buffer): ZipEntry[] {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 0xffff); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not a zip file (no end of central directory)");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const entries: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("corrupt zip central directory");
    const method = buf.readUInt16LE(p + 10);
    const compressedSize = buf.readUInt32LE(p + 20);
    const nameLength = buf.readUInt16LE(p + 28);
    const extraLength = buf.readUInt16LE(p + 30);
    const commentLength = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf-8", p + 46, p + 46 + nameLength);
    entries.push({
      name,
      read() {
        const dataStart = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28);
        const data = buf.subarray(dataStart, dataStart + compressedSize);
        if (method === 0) return Buffer.from(data);
        if (method === 8) return zlib.inflateRawSync(data);
        throw new Error(`${name}: unsupported zip compression method ${method}`);
      },
    });
    p += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** Content-gate findings from the package parts. */
function inspectPackage(buf: Buffer): Map<GateItem, string[]> {
  const found = new Map<GateItem, string[]>();
  const add = (item: GateItem, detail: string) => found.set(item, [...(found.get(item) ?? []), detail]);
  const entries = readZip(buf);
  for (const entry of entries) {
    const name = entry.name.toLowerCase();
    if (name.endsWith("vbaproject.bin") || name.endsWith("vbadata.xml")) add("vba", entry.name);
    if (name.startsWith("xl/externallinks/") && name.endsWith(".xml")) add("external-links", entry.name);
    if (name.endsWith(".rels")) {
      const xml = entry.read().toString("utf-8");
      for (const rel of xml.match(/<Relationship\b[^>]*>/g) ?? []) {
        if (!/TargetMode="External"/i.test(rel)) continue;
        const target = rel.match(/Target="([^"]*)"/)?.[1] ?? "";
        const type = rel.match(/Type="([^"]*)"/)?.[1] ?? "";
        if (/\/hyperlink$/.test(type)) {
          if (!/^https:/i.test(target)) add("http-links", `${target} (${entry.name})`);
        } else if (/externalLink/i.test(type) || /externalLinkPath/i.test(type)) {
          add("external-links", `${target} (${entry.name})`);
        }
      }
    }
  }
  return found;
}

// ---------------------------------------------------------------------------------------------
// Workbook-level inspection with the SDK
// ---------------------------------------------------------------------------------------------

interface WorkbookFindings {
  gate: Map<GateItem, string[]>;
  nameErrors: string[];
  otherErrors: string[];
  sheets: number;
  formulas: number;
}

async function inspectWorkbook(wb: IWorkbook): Promise<WorkbookFindings> {
  const gate = new Map<GateItem, string[]>();
  const add = (item: GateItem, detail: string) => gate.set(item, [...(gate.get(item) ?? []), detail]);
  const nameErrors: string[] = [];
  const otherErrors: string[] = [];
  let formulas = 0;

  // The 0.8.0-beta.1 xlsx writer does not persist script modules, so today only an .sxl can carry
  // them.
  const modules = wb.getScripts().getModules().getItems();
  for (const module of modules) add("scripts", (module as any).getName?.() ?? "module");

  const externalBooks = (wb.getExternalBooks().toJSON() ?? []).length;
  if (externalBooks > 0) add("external-links", `${externalBooks} external workbook reference(s)`);

  for (const named of wb.getNames().getItems()) {
    const text = JSON.stringify(named.toJSON?.() ?? "");
    if (FETCH_REGEX.test(text)) add("fetch", `defined name ${named.getName()}`);
  }

  const sheets = wb.getSheets().getItems();
  for (const sheet of sheets) {
    // A full recalculation, so the values checked are the engine's, not whatever the file cached.
    await sheet.calculate(true);
  }
  await wb.getCalculation().ready();

  for (const sheet of sheets) {
    const at = (address: string) => `${sheet.getName()}!${address}`;

    for (const region of sheet.getRegions("formula")) {
      formulas += region.getRowCount() * region.getColumnCount();
      const text = region.getCell().getFormulaText() ?? "";
      const fetch = text.match(FETCH_REGEX);
      if (fetch) add("fetch", `${at(region.getAddress())} uses ${fetch[1].toUpperCase()}`);
      if (/HYPERLINK\s*\(/i.test(text) && HTTP_LITERAL_REGEX.test(text)) {
        add("http-links", `${at(region.getAddress())} HYPERLINK to a non-https address`);
      }
    }

    for (const region of sheet.getRegions("error")) {
      for (const row of region.getRange().getCells()) {
        for (const cell of row) {
          const text = cell.getText();
          if (text === "#NAME?") nameErrors.push(at(cell.getAddress()));
          else if (text.startsWith("#")) otherErrors.push(`${at(cell.getAddress())} ${text}`);
        }
      }
    }

    for (const region of sheet.getRegions("hyperlink")) {
      for (const row of region.getRange().getCells()) {
        for (const cell of row) {
          const address = cell.getHyperlink()?.getAddress() ?? "";
          if (!address || address.startsWith("#") || /^https:/i.test(address)) continue;
          add("http-links", `${at(cell.getAddress())} -> ${address}`);
        }
      }
    }
  }

  return { gate, nameErrors, otherErrors, sheets: sheets.length, formulas };
}

// ---------------------------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------------------------

async function validateTemplate(file: string): Promise<string> {
  const where = posix(path.relative(process.cwd(), file));
  const buf = await fs.readFile(file);
  const format = TEMPLATE_FORMATS[path.extname(file).toLowerCase()];

  const gate = new Map<GateItem, string[]>();
  const merge = (from: Map<GateItem, string[]>) => {
    for (const [item, details] of from) gate.set(item, [...(gate.get(item) ?? []), ...details]);
  };
  if (buf.byteLength > MAX_BYTES) gate.set("size", [`${(buf.byteLength / 1024 / 1024).toFixed(2)} MB`]);
  if (format === "xlsx") {
    try {
      merge(inspectPackage(buf));
    } catch (err: any) {
      fail(where, `cannot read the package: ${err.message}`);
      return "unreadable";
    }
  }

  let wb: IWorkbook | null = null;
  try {
    wb = await WorkbookIO.read({
      source: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      name: path.basename(file),
      format,
    } as any);
  } catch (err: any) {
    fail(where, `does not open: ${err.message ?? err}`);
    return "does not open";
  }
  if (!wb) {
    fail(where, "does not open");
    return "does not open";
  }

  try {
    await wb.getCalculation().start();
    const findings = await inspectWorkbook(wb);
    merge(findings.gate);
    if (findings.nameErrors.length) {
      fail(where, `#NAME? in ${findings.nameErrors.length} cell(s): ${findings.nameErrors.slice(0, 10).join(", ")}`);
    }
    for (const [item, details] of gate) {
      fail(where, `contains ${GATE[item]} (${details.slice(0, 5).join("; ")})`);
    }
    const warnings = findings.otherErrors.length ? `, ${findings.otherErrors.length} other error value(s)` : "";
    return `${findings.sheets} sheet(s), ${findings.formulas} formula cell(s), ${(buf.byteLength / 1024).toFixed(1)} KB${warnings}`;
  } finally {
    await wb.getCalculation().ready();
    await wb.getCalculation().stop();
    wb.close();
  }
}

async function build(): Promise<void> {
  const categories: Category[] = [];
  const entries: Entry[] = [];
  const copies: { from: string; to: string }[] = [];

  const top = await fs.readdir(SRC, { withFileTypes: true });
  for (const item of top) {
    if (item.name.startsWith(".")) continue;
    if (item.isFile()) {
      fail(`templates/${item.name}`, "ignored; templates must sit in a category folder (templates/<category>/<name>.xlsx|.sxl|.csv)");
    }
  }

  for (const dir of top.filter((d) => d.isDirectory() && !d.name.startsWith(".")).sort((a, b) => a.name.localeCompare(b.name))) {
    const id = dir.name;
    const catDir = path.join(SRC, id);
    const catBefore = problems.length;
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(id)) fail(`templates/${id}`, "category folders must be lowercase-kebab-case");

    const files = await fs.readdir(catDir, { withFileTypes: true });
    const names = new Set(files.filter((f) => f.isFile()).map((f) => f.name));
    const catInfo = names.has("_category.json") ? await readCategory(path.join(catDir, "_category.json")) : {};
    if (problems.length > catBefore) {
      console.log(`SKIP  ${id}/  category left out`);
      continue;
    }
    /** Template file names by base name. Thumbnails and descriptions are matched on the base name. */
    const templates = new Map<string, string[]>();
    for (const name of names) {
      const ext = path.extname(name).toLowerCase();
      if (!TEMPLATE_EXTENSIONS.includes(ext) || name.startsWith("~$")) continue;
      const base = name.slice(0, name.length - ext.length);
      templates.set(base, [...(templates.get(base) ?? []), name]);
    }

    let count = 0;
    for (const f of files) {
      const rel = `templates/${id}/${f.name}`;
      // Dot files, and the `~$<name>.xlsx` lock file Excel keeps beside a workbook it has open.
      if (f.name.startsWith(".") || f.name.startsWith("~$")) continue;
      if (f.isDirectory()) {
        fail(rel, "ignored; nested folders are not supported, one category level only");
        continue;
      }
      const ext = path.extname(f.name).toLowerCase();
      const base = f.name.slice(0, f.name.length - ext.length);
      if (f.name === "_category.json") continue;
      if (THUMBNAIL_EXTENSIONS.includes(ext)) {
        const owner = base.endsWith(DARK_SUFFIX) ? base.slice(0, -DARK_SUFFIX.length) : base;
        if (!templates.has(owner)) fail(rel, `ignored; thumbnail has no matching template ${owner}.xlsx|.sxl|.csv`);
        continue;
      }
      if (ext === DESCRIPTION_EXTENSION) {
        if (!templates.has(base)) fail(rel, `ignored; description has no matching template ${base}.xlsx|.sxl|.csv`);
        continue;
      }
      if (!TEMPLATE_EXTENSIONS.includes(ext)) {
        fail(rel, `ignored; unsupported file type "${ext}", templates are ${TEMPLATE_EXTENSIONS.join(", ")}`);
        continue;
      }
      const before = problems.length;
      if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(base)) fail(rel, "template file names must be lowercase-kebab-case");
      if (templates.get(base)!.length > 1) fail(rel, `another template shares the name "${base}" (${templates.get(base)!.join(", ")})`);

      let summary: string;
      try {
        summary = await validateTemplate(path.join(catDir, f.name));
      } catch (err: any) {
        fail(rel, `validation threw: ${err?.message ?? err}`);
        summary = "validation threw";
      }
      const descriptionName = `${base}${DESCRIPTION_EXTENSION}`;
      const description = names.has(descriptionName) ? await readDescription(path.join(catDir, descriptionName)) : undefined;

      if (problems.length > before) {
        console.log(`SKIP  ${id}/${f.name}  ${summary}`);
        continue;
      }
      console.log(`  ok  ${id}/${f.name}  ${summary}`);

      const entry: Entry = {
        path: `${id}/${f.name}`,
        category: id,
        title: toName(base),
      };
      if (description) entry.description = description;
      for (const [key, suffix] of [["thumbnail", ""], ["thumbnailDark", DARK_SUFFIX]] as const) {
        const handMade = THUMBNAIL_EXTENSIONS.map((e) => `${base}${suffix}${e}`).find((n) => names.has(n));
        const rendered = path.join(RENDERED, id, `${base}${suffix}.webp`);
        if (handMade) {
          entry[key] = `${id}/${handMade}`;
          copies.push({ from: path.join(catDir, handMade), to: path.join(OUT, id, handMade) });
        } else if (await exists(rendered)) {
          entry[key] = `${id}/${base}${suffix}.webp`;
          copies.push({ from: rendered, to: path.join(OUT, id, `${base}${suffix}.webp`) });
        }
      }
      copies.push({ from: path.join(catDir, f.name), to: path.join(OUT, id, f.name) });
      entries.push(entry);
      count++;
    }

    if (count === 0) {
      fail(`templates/${id}`, "category left out; it has no valid templates");
      continue;
    }
    const category: Category = { id, title: catInfo.title ?? toName(id) };
    if (catInfo.description) category.description = catInfo.description;
    if (catInfo.icon) category.icon = catInfo.icon;
    if (catInfo.order !== undefined) category.order = catInfo.order;
    categories.push(category);
  }

  if (problems.length) {
    console.warn(`\n${problems.length} problem(s):`);
    for (const { where, message } of problems) {
      console.warn(`  - ${where}: ${message}`);
      // A GitHub Actions annotation, so a left-out template shows on the run summary and the PR.
      if (process.env.GITHUB_ACTIONS) console.log(`::warning file=${where}::${message.replace(/\r?\n/g, " ")}`);
    }
  }
  if (entries.length === 0) {
    console.error("\nBuild failed: no valid templates found");
    process.exit(1);
  }
  if (STRICT && problems.length) {
    console.error(`\nBuild failed (--strict) with ${problems.length} problem(s)`);
    process.exit(1);
  }

  categories.sort((a, b) => (a.order ?? Infinity) - (b.order ?? Infinity) || a.title.localeCompare(b.title));
  const rank = new Map(categories.map((c, i) => [c.id, i]));
  entries.sort((a, b) => rank.get(a.category)! - rank.get(b.category)! || a.title.localeCompare(b.title));

  for (const { from, to } of copies) {
    await fs.mkdir(path.dirname(to), { recursive: true });
    await fs.copyFile(from, to);
  }

  const catalog = {
    kind: "templates",
    package: PACKAGE,
    meta: {
      schema: SCHEMA,
      categories: categories.map(({ order, ...rest }) => rest),
    },
    entries,
  };
  await fs.writeFile(path.join(OUT, "contents.json"), JSON.stringify(catalog, null, 2) + "\n");

  // The root listing: every file and folder under dist/, except itself.
  const files: string[] = [];
  const folders: string[] = [];
  async function walk(dir: string): Promise<void> {
    for (const d of (await fs.readdir(dir, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name))) {
      const full = path.join(dir, d.name);
      const rel = posix(path.relative(DIST, full));
      if (d.isDirectory()) {
        folders.push(rel);
        await walk(full);
      } else if (rel !== "contents.json") {
        files.push(rel);
      }
    }
  }
  await walk(DIST);
  await fs.writeFile(
    path.join(DIST, "contents.json"),
    JSON.stringify({ entries: files.map((p) => ({ path: p })), folders }, null, 2) + "\n",
  );

  console.log(`\nBuilt ${entries.length} template(s) in ${categories.length} categor${categories.length === 1 ? "y" : "ies"}`);
  console.log(`- dist/contents.json (root listing, ${files.length} files)`);
  console.log(`- dist/${PACKAGE}/contents.json (catalog)`);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
