# Workbook Templates

![SheetXL](https://www.sheetxl.com/logo-text.svg)

[![License: CC0-1.0](https://img.shields.io/badge/license-CC0--1.0-blue.svg)](LICENSE)
[![Discord](https://img.shields.io/discord/1141404921246257223)](https://discord.gg/NTKdwUgK9p)

## Overview

The workbook templates that appear in the SheetXL **New Workbook** gallery. The repo is published to
npm as [`@sheetxl/workbook-templates`](https://www.npmjs.com/package/@sheetxl/workbook-templates),
and the SheetXL studio reads the catalog straight from unpkg:

```text
https://unpkg.com/@sheetxl/workbook-templates/dist/workbook-templates/contents.json
```

Every template is original content, dedicated to the public domain under [CC0 1.0](LICENSE). Use
them however you like.

## Adding a template

Drop an `.xlsx`, `.sxl` or `.csv` into a category folder under `templates/` and open a pull request.
That is all. CI validates the workbook, builds the index and publishes it.

- **The file name is the title.** `yearly-calendar.xlsx` is shown as "Yearly Calendar". Use
  lowercase-kebab-case names.
- **The folder is the category.** Use an existing folder, or create a new one. There is one level
  only; a template directly in `templates/` is ignored.
- **A `<name>.txt` beside the template is its description.** The gallery shows it as the tile's
  tooltip and searches it. Keep it to a sentence or two of plain text (under 300 characters); line
  breaks are folded to spaces, and markup keeps the template out of the catalog.
- **Thumbnails come from CI.** You do not need to make one. CI opens each template in the latest beta
  SheetXL studio (`render/`) and photographs its first screen twice: once light, and once dark with
  the dark grid on. The headers and the selection are left out. A hand-made `<name>.png` (or
  `<name>.dark.png`) beside the template is used instead, for a template whose first screen is not
  its best.

To check a template before you push, run `npm install` and `npm run build -- --strict`. To see its thumbnails,
run `npx playwright install chromium` once, then `npm run thumbnails` before the build; the images
land in `.thumbnails/`.

Templates must be your own work, or carry a license that allows redistribution under CC0. Do not
submit Microsoft Office templates or close copies of them.

## Layout

```text
templates/
  <category>/
    _category.json        optional: { "title", "description", "order", "icon" }
    <name>.xlsx           a template (or <name>.sxl, <name>.csv)
    <name>.txt            optional plain-text description, the tile's tooltip
    <name>.png            optional hand-made thumbnail
    <name>.dark.png       optional hand-made dark thumbnail
config/thumbnails.ts      renders .thumbnails/<category>/<name>[.dark].webp for each template
config/build.ts           validates every template, then writes dist/
render/                   the page that renders one template for its thumbnail
```

A category folder's `_category.json` sets its display title, description, sort `order` and icon.
Without one, the title comes from the folder name ("project-management" becomes "Project
Management"). Categories sort by `order`, then by title. Templates sort by title within a category.

The build writes:

```text
dist/contents.json                                root listing of every file and folder
dist/workbook-templates/contents.json             the catalog
dist/workbook-templates/<category>/<name>.xlsx      or .sxl, .csv
dist/workbook-templates/<category>/<name>.webp    the rendered thumbnail, or a hand-made .png
dist/workbook-templates/<category>/<name>.dark.webp   the same on the dark grid
```

The catalog looks like this. Entry paths are relative to `dist/workbook-templates/`.

```jsonc
{
  "kind": "templates",
  "package": "workbook-templates",
  "meta": {
    "schema": 1,
    "categories": [{ "id": "finance", "title": "Finance", "description": "..." }]
  },
  "entries": [
    {
      "path": "finance/loan-calculator.xlsx", "category": "finance", "title": "Loan Calculator",
      "description": "Enter the loan amount, rate and term to get the payment, ...",
      "thumbnail": "finance/loan-calculator.webp", "thumbnailDark": "finance/loan-calculator.dark.webp"
    }
  ]
}
```

## What the build checks

Each template must open in the SheetXL SDK and fully recalculate with no `#NAME?` in any cell. A
template that also contains any of these fails the checks:

- SheetXL script modules;
- a VBA project;
- links to other workbooks;
- `WEBSERVICE`, `FILTERXML`, `IMPORTDATA` or another function that fetches from the network;
- a hyperlink that is not `https:`;
- more than 2 MB.

Two templates with the same name and different extensions also fail, as does an invalid
description. Any other file in a category folder, besides `_category.json`, thumbnails and
descriptions, is ignored with a warning.

A template that fails is left out of the catalog, and a category folder with an invalid name or
`_category.json` is left out with its templates. The build lists every problem as a warning (and as
an annotation in GitHub Actions) and publishes the rest; it fails only when no template passes.
`npm run build -- --strict` fails on any problem instead.

## Publishing

A push to `main` that changes `templates/` or `config/` builds the catalog, bumps the patch version,
tags it, publishes to npm and creates a GitHub release. To publish without a template change, run
the workflow from the Actions tab (**Run workflow**), or put `[publish]` in the message of a commit
that changes `render/`, `package.json` or the workflow.

## Additional Resources

- **[Main GitHub](https://github.com/sheetxl)**
- **[Discord](https://discord.gg/NTKdwUgK9p)**
- **[Website](https://www.sheetxl.com)**
- **[Developer Docs](https://www.sheetxl.com/docs)**

## License

The templates are dedicated to the public domain under [CC0 1.0 Universal](LICENSE).
