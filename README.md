# TMR Tracker → UpLook Sheet Builder

Static page that turns a monthly TMR Tracker export into the ingredient sheet UpLook accepts.
Premixes (PREMIX, 2PRMIX) are broken out into their raw ingredients, and the Bovaer Premix
(BOVAR) is split into Bovaer and ground corn. Everything runs in the browser, so no feed data
leaves the page.

Live: https://pluceyelancoah.github.io/tmr-uplook-builder/ (GitHub Pages, `main` branch, repo root).
**Pushing to `main` deploys straight away.**

## Supported reports
| Report | Format | Notes |
|---|---|---|
| **Load report** | CSV | Preferred. Rations that contain PREMIX/2PRMIX count as milking cows. The Bovaer share of BOVAR is read from that month's BOVAR batches. |
| Ingredient report (older) | Excel or derived CSV | Bovaer share comes from Advanced settings (default 7.76% DM). |
| Ingredient/Premix Usage | CSV | Rejected: it double counts the premix mixing loads and doesn't show what's in each premix. |

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
