# TapFeed / TMR Tracker → UpLook Sheet Builder

Static page that turns a monthly TapFeed Load report (or the older TMR Tracker ingredient
report) into the ingredient sheet UpLook accepts.
Premixes (PREMIX, 2PRMIX) are broken out into their raw ingredients, and then the Bovaer
Premix (BOVAR) is broken out into Bovaer and ground corn. Each one uses its own batch recipe
from the same report; there are no built-in defaults. If a premix's batches are missing, the
page warns instead of guessing. The page also shows the Bovaer dose (mg 3-NOP per kg dry
matter, target 60), calculated from Bovaer as fed × 10% ÷ dry matter eaten. Everything runs
in the browser, so no feed data leaves the page.

Downloading the UpLook sheet also downloads a **Bovaer Dose Confirmation Report (PDF)** in the same
layout as the Bovaer Dose Calculator (COM-BCAL). Each premix is a "Bovaer in Blend, nested in premix"
source, all as-fed, with the nested % set so the calculator math reproduces the dose above exactly.
Typing those inputs into the calculator gives the same dose.

Live: https://pluceyelancoah.github.io/tmr-uplook-builder/ (GitHub Pages, `main` branch, repo root).
**Pushing to `main` deploys straight away.**

## Supported reports
Several files can be dropped at once. Each is identified, marked used / needs another file / not needed / can't be used,
and the page says what to add. Priority: Load report > TMR Tracker report > Ingredient usage + Premixes.

| Report | Format | Notes |
|---|---|---|
| **TapFeed Load report** | CSV | Preferred, the only file needed; unfiltered is fine. Rations that contain PREMIX/2PRMIX count as milking cows. Single-month method, so a verifier can redo it from that month's report alone: Bovaer fed = PREMIX/2PRMIX fed (as fed) × Bovaer premix % of the premix batches made that month × Bovaer % of the Bovaer premix batches made that month. If the file covers more than one month, only the last is counted. A recipe with no batch in the month uses the last batch before it, with a note. Also gives the recipe (intended) dose from the batch call weights. |
| TapFeed Ingredient usage + Premixes | CSV + PDF | Together only: recipe (intended) dose from what each pen was fed and the Premixes recipe. Used as a cross-check when a Load report is also given. The usage report has no date range. |
| TMR Tracker ingredient report (older) | Excel or derived CSV | Only works if the export includes the BOVAR batch section. Outside that section, the code `Bovaer` means the BOVAR blend. |
| TapFeed Ingredient/Premix Usage | CSV | Can't be used: it double counts the premix mixing loads and doesn't show what's in each premix. |
| TapFeed Recipes | PDF | Not needed. |

## Layout
- `index.html`: markup only
- `js/core.js`: parsing and calculation, no DOM (also used by the tests)
- `js/app.js`: page wiring and rendering
- `js/pdf.js`: dose report PDF, ported from COM-BCAL `src/web/src/utils/pdf-export.ts`
- `js/logo.js`: Bovaer logo as a data URL (the page CSP blocks fetching images for the PDF)
- `css/input.css` → `css/app.css`: Tailwind source → built output (committed)
- `vendor/xlsx.full.min.js`, `vendor/jspdf.umd.min.js`: SheetJS and jsPDF, copied from `node_modules` (committed)
- `vendor/pdf.min.js`, `vendor/pdf.worker.min.js`: pdf.js 3.x, loaded only when a PDF is added; the worker script runs on the page (no Worker), so it works from `file://` and under the CSP
- `test/pdf-lines.js`: prints the text lines pdf.js extracts from a PDF (for working on the Premixes parser)

## Develop
```powershell
npm install
npm run build        # rebuild css/app.css and vendor/ — commit both
npm run watch:css    # while editing markup/classes
npm run serve        # http://localhost:8080
npm test
```

## Tests
- `test/core.test.js`: made-up reports, safe to publish.
- `test/samples.test.js`: real farm reports in `raw/`, compared against `raw/*.expected.json`.
  `raw/` is git-ignored because it contains farm data and this repo is public. After an intended
  change in output, run `npm run snapshot`, then review the change before accepting it.

CI runs the tests and fails if `css/app.css` or `vendor/` weren't rebuilt.
