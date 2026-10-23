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
| Report | Format | Notes |
|---|---|---|
| **TapFeed Load report** | CSV | Preferred. Rations that contain PREMIX/2PRMIX count as milking cows. Worked batch by batch: each BOVAR batch's recipe applies until the next one is made, and premix batches are fed out first-in, first-out. If the export starts before the month (ideally at the last BOVAR batch before it), only the last month is counted and earlier rows are batch history; otherwise premix mixed before the month's first BOVAR batch borrows that batch's recipe, with a note. |
| TMR Tracker ingredient report (older) | Excel or derived CSV | Only works if the export includes the BOVAR batch section. Outside that section, the code `Bovaer` means the BOVAR blend. |
| TapFeed Ingredient/Premix Usage | CSV | Rejected: it double counts the premix mixing loads and doesn't show what's in each premix. |

## Layout
- `index.html`: markup only
- `js/core.js`: parsing and calculation, no DOM (also used by the tests)
- `js/app.js`: page wiring and rendering
- `js/pdf.js`: dose report PDF, ported from COM-BCAL `src/web/src/utils/pdf-export.ts`
- `js/logo.js`: Bovaer logo as a data URL (the page CSP blocks fetching images for the PDF)
- `css/input.css` → `css/app.css`: Tailwind source → built output (committed)
- `vendor/xlsx.full.min.js`, `vendor/jspdf.umd.min.js`: SheetJS and jsPDF, copied from `node_modules` (committed)

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
