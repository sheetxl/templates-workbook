/**
 * Renders two thumbnails for every template, light and dark, into
 * `.thumbnails/<category>/<name>.webp` and `<name>.dark.webp`, which `config/build.ts` then
 * publishes beside the workbook.
 *
 *   npm run thumbnails [-- <filter>]
 *
 * Each template is opened by `render/` (the pinned @sheetxl/studio-mui) in headless Chromium,
 * photographed once the grid has painted and its fonts have loaded, and scaled down to webp. The
 * dark one is the dark app with the dark grid on. A hand-made `templates/<category>/<name>.png|.webp`
 * (or `<name>.dark.png|.webp`) wins over the rendered one for that scheme, so it is skipped.
 *
 * Needs a Chromium for Playwright: `npx playwright install chromium` (CI adds `--with-deps`).
 */
import fs from "node:fs/promises";
import path from "node:path";

import { build, preview } from "vite";
import { chromium } from "playwright";
import sharp from "sharp";

const SRC = path.resolve("templates");
const OUT = path.resolve(".thumbnails");
const RENDER_CONFIG = path.resolve("render/vite.config.ts");

/** The rendered viewport, in CSS pixels. */
const VIEWPORT = { width: 560, height: 350 };
const SCALE = 2;
/** The grid's cell area: every pane, frozen ones included, and none of the headers. */
const CELLS_SELECTOR = ".sheetxl-sheet-primary";
/**
 * What the grid draws around the cells inside its cell area, in CSS pixels, measured on the pinned
 * studio: a 3px header-colored strip on the left, a 1px line on top and 2px on the right. Not
 * photographed.
 */
const CELLS_INSET = { left: 3, top: 1, right: 2, bottom: 0 };
/** The published thumbnail, in pixels. Cropped from the top left of the cell area to 8:5. */
const THUMBNAIL = { width: 480, height: 300 };
const LOAD_TIMEOUT = 60_000;

const HAND_MADE = [".png", ".webp"];
const TEMPLATE_REGEX = /\.(?:xlsx|sxl|csv)$/i;

/** Each scheme's file-name suffix, before the extension. */
const SCHEMES = { light: "", dark: ".dark" } as const;
type Scheme = keyof typeof SCHEMES;

interface Job {
  rel: string;
  scheme: Scheme;
}

async function listJobs(filter?: string): Promise<Job[]> {
  const found: Job[] = [];
  for (const dir of await fs.readdir(SRC, { withFileTypes: true })) {
    if (!dir.isDirectory() || dir.name.startsWith(".")) continue;
    const names = new Set(await fs.readdir(path.join(SRC, dir.name)));
    for (const name of names) {
      // `~$<name>.xlsx` is the lock file Excel keeps beside a workbook it has open.
      if (!TEMPLATE_REGEX.test(name) || name.startsWith("~$")) continue;
      const base = name.replace(TEMPLATE_REGEX, "");
      const rel = `${dir.name}/${name}`;
      if (filter && !rel.includes(filter)) continue;
      for (const [scheme, suffix] of Object.entries(SCHEMES) as [Scheme, string][]) {
        if (HAND_MADE.some((ext) => names.has(base + suffix + ext))) continue;
        found.push({ rel, scheme });
      }
    }
  }
  return found.sort((a, b) => a.rel.localeCompare(b.rel) || a.scheme.localeCompare(b.scheme));
}

async function main(): Promise<void> {
  const jobs = await listJobs(process.argv[2]);
  if (jobs.length === 0) {
    console.log("No templates need a rendered thumbnail.");
    return;
  }

  await build({ configFile: RENDER_CONFIG });
  const server = await preview({ configFile: RENDER_CONFIG, preview: { port: 0, host: "127.0.0.1" } });
  const origin = server.resolvedUrls?.local[0];
  if (!origin) throw new Error("The render page did not start.");

  const browser = await chromium.launch();
  const failures: string[] = [];
  try {
    const context = await browser.newContext({ viewport: VIEWPORT, deviceScaleFactor: SCALE });
    for (const { rel, scheme } of jobs) {
      const label = `${rel} (${scheme})`;
      const page = await context.newPage();
      const errors: string[] = [];
      page.on("pageerror", (e) => errors.push(e.message));
      try {
        await page.goto(`${origin}?src=${encodeURIComponent(rel)}&scheme=${scheme}`);
        await page.waitForFunction(() => window.__thumbnail !== undefined, undefined, { timeout: LOAD_TIMEOUT });
        const outcome = await page.evaluate(() => window.__thumbnail);
        if (outcome !== "ready") throw new Error(outcome?.error ?? "did not load");
        await page.evaluate(async () => {
          await document.fonts.ready;
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        });
        const cells = await page.locator(CELLS_SELECTOR).first().boundingBox();
        if (!cells) throw new Error(`no ${CELLS_SELECTOR} on the page`);
        const png = await page.screenshot({
          type: "png",
          clip: {
            x: cells.x + CELLS_INSET.left,
            y: cells.y + CELLS_INSET.top,
            width: cells.width - CELLS_INSET.left - CELLS_INSET.right,
            height: cells.height - CELLS_INSET.top - CELLS_INSET.bottom,
          },
        });
        const webp = await sharp(png).resize(THUMBNAIL.width, THUMBNAIL.height, { fit: "cover", position: "left top" }).webp({ quality: 82 }).toBuffer();
        const out = path.join(OUT, rel.replace(TEMPLATE_REGEX, `${SCHEMES[scheme]}.webp`));
        await fs.mkdir(path.dirname(out), { recursive: true });
        await fs.writeFile(out, webp);
        console.log(`  ok  ${label}  ${(webp.length / 1024).toFixed(1)} KB`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        failures.push(`${label}: ${message}${errors.length ? ` (${errors.join("; ")})` : ""}`);
        console.log(`FAIL  ${label}  ${message}`);
      } finally {
        await page.close();
      }
    }
  } finally {
    await browser.close();
    await server.close();
  }

  if (failures.length) {
    console.error(`\n${failures.length} thumbnail(s) failed:`);
    for (const f of failures) console.error(`  - ${f}`);
    process.exit(1);
  }
  console.log(`\nRendered ${jobs.length} thumbnail(s) into .thumbnails/`);
}

declare global {
  interface Window { __thumbnail?: "ready" | { error: string } }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
