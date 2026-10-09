# TapFeed / TMR Tracker → UpLook Sheet Builder

Static page that turns a monthly TapFeed Load report (or the older TMR Tracker ingredient
report) into the ingredient sheet UpLook accepts.
Premixes (PREMIX, 2PRMIX) are broken out into their raw ingredients, and then the Bovaer
Premix (BOVAR) is broken out into Bovaer and ground corn. Each one uses its own batch recipe
from the same report; there are no built-in defaults. If a premix's batches are missing, the
page warns instead of guessing. The page also shows the Bovaer dose (mg 3-NOP per kg dry
matter, target 60), calculated from Bovaer as fed × 10% ÷ dry matter eaten. Everything runs
in the browser, so no feed data leaves the page.

Live: https://pluceyelancoah.github.io/tmr-uplook-builder/ (GitHub Pages, `main` branch, repo root).
**Pushing to `main` deploys straight away.**

## Supported reports
| Report | Format | Notes |
|---|---|---|
| **TapFeed Load report** | CSV | Preferred. Rations that contain PREMIX/2PRMIX count as milking cows. A month with no BOVAR batch made in it can't be broken out. |
| TMR Tracker ingredient report (older) | Excel or derived CSV | Only works if the export includes the BOVAR batch section. Outside that section, the code `Bovaer` means the BOVAR blend. |
| TapFeed Ingredient/Premix Usage | CSV | Rejected: it double counts the premix mixing loads and doesn't show what's in each premix. |

## Layout
- `index.html`: markup only
- `js/core.js`: parsing and calculation, no DOM (also used by the tests)
- `js/app.js`: page wiring and rendering
- `css/input.css` → `css/app.css`: Tailwind source → built output (committed)
- `vendor/xlsx.full.min.js`: SheetJS, copied from `node_modules` (committed)

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
