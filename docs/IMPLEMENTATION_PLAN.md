# Fygaro Catalog Automation - Chrome Extension

## Context

Ana Elisa Villalaz has a catalog of 699 services in `docs/Catálogo de Productos y Servicios Fygaro.xlsx`
(sheet `"Logros "`, note the trailing space). Every service needs to exist as a Product in the Fygaro
merchant app, and every product needs a Fygaro payment link, and that link has to be written back into
column H of the spreadsheet next to the right row.

Doing this by hand is 699 repetitions of a 7 page workflow. This extension automates that loop inside the
user's own logged in browser session, with visible progress, pause and resume, and a patched spreadsheet
at the end.

Confirmed decisions:
- Local `.xlsx` in and out. No Google API, no OAuth. Parsing and rewriting is done in pure vanilla JS.
- Rows that already have a value in column H are skipped. This is what makes resume safe.
- Advanced Options DOM is now captured, so the three checkboxes are targeted by exact `name` attributes.
- A failed step pauses the run and asks the user (Retry / Skip row / Stop). Nothing fails silently.
- Light theme only, with a deliberate professional design system. No dark mode, no OS theme following.

## Constraints

- Vanilla JavaScript only. No React, no framework, no bundler, no build step.
- Manifest V3, all scripts classic (non module), sharing one global `FYG` namespace so the same shared
  files load in the service worker (`importScripts`), the content script (manifest `js` array) and the
  side panel (`<script src>`).
- Everything installed project local only. Target: zero runtime dependencies.
- No em dash anywhere in the project.
- Git: local commits only, never push. First commit message exactly
  `Muhammad Abdullah Awais (www.abdullahawais.com)`, no co-author trailer. Git identity is already set
  globally to the correct name and email, so no git config changes are needed.

## Data facts established from the real file

| Column | Header | Maps to |
|---|---|---|
| C | Código | Product `code`, and the Fygaro Link `name` |
| D | Servicios | Product `name` |
| E | Precio Total | Product `price` |
| H | Link | Written back after the link is created |

- 699 data rows, rows 2 to 700. Table range is `A1:H1100`, sheet XML has no `<dimension>` element.
- Price formats are mixed and both must parse:
  - `B/.625,00` (450 rows) and `B/.25,00` (6 rows) -> comma is the decimal separator
  - `B/.1.125,00` (210 rows) -> dot is thousands, comma is decimal
  - `B/.100.00` (27 rows) and `B/.80.00` (6 rows) -> dot is the decimal separator
  - Rule: strip the leading `B/.`, then if a comma is present treat it as the decimal point and drop all
    dots, otherwise treat the final dot as the decimal point.
- Only row 18 already has a Link, so 698 rows are in scope on a fresh run.
- Duplicate data exists: 2 duplicate Códigos and 1 duplicate Servicios name. Product lookup therefore
  matches on name plus code plus price, and takes the first match, which is correct because the products
  list is ordered newest first.
- The workbook is currently open in Excel. The extension never writes to the original file, it produces a
  download, so this is not a blocker, but the user should close Excel before replacing the file.

## Architecture

```
manifest.json                       MV3, side panel, host_permissions https://www.fygaro.com/*
icons/                              16 / 32 / 48 / 128 png, generated locally
src/shared/util.js                  sleep, randomDelay, clamp, csv, safe log ring buffer
src/shared/price.js                 parsePrice covering both number formats
src/shared/zip.js                   ZIP reader and writer on DecompressionStream / CompressionStream + CRC32
src/shared/xlsx-read.js             workbook.xml + rels + sharedStrings + sheet XML -> rows
src/shared/xlsx-write.js            inject column H cells into sheet XML, rebuild the zip
src/shared/state.js                 chrome.storage.local schema, get/patch, message name constants
src/background/service-worker.js    single owner of run state, messaging hub, badge, notifications
src/content/dom.js                  waitFor, isVisible, findByText, setNativeValue, humanClick
src/content/steps.js                the seven step handlers
src/content/main.js                 SPA route watcher, tick loop, action lock
src/sidepanel/sidepanel.html|css|js the whole UI
tests/selectors.test.html           runs the seven step locators against all eight saved DOM snapshots
tests/xlsx.test.html                round trips the real xlsx: read -> patch H -> read back
tests/price.test.mjs                node --test, zero deps, all five price patterns
README.md                           professional docs including developer details
```

### Why side panel and not popup

The run takes many hours. A popup closes the moment focus moves, which destroys the UI. `chrome.sidePanel`
stays open next to the Fygaro tab for the whole run. State lives in `chrome.storage.local` regardless, so
even a full browser restart resumes.

### Selector strategy (important)

Fygaro uses hashed CSS module class names such as `tJCkqVa46nRptzcAqFwx`. Those break on any deploy.
Priority order for every locator:

1. `name` attribute: `input[name="name"]`, `input[name="code"]`, `input[name="price"]`,
   `select[name="currency"]`, `select[name="product_type"]`, `select[name="show_in_website"]`,
   `input[name="require_phone"]`, `input[name="require_legal_id"]`, `input[name="require_billing_address"]`
2. `href`: sidebar Products is `a[href="/en/app/products/"]`, top right Create is the
   `a[href="/en/app/products/add/"]` that is **not** inside the empty state block, Back to Home is
   `a[href="/en/app/dashboard/"]`
3. Visible text: `Save`, `Create Fygaro Link for Product`, `Advanced Options`, matched case insensitively
   against EN and ES variants
4. Hashed class as a last resort fallback only

One trap the captured DOM revealed: `name="max_successful_payments"` exists **twice** on the link form,
once as a checkbox in General Settings and once as a number input in Advanced Options. Every input locator
is therefore type qualified, for example `input[type="checkbox"][name="..."]`.

`isVisible()` filters out the responsive duplicate of a control (the product form renders the Currency /
Type of Item / Show In Website group twice, once per layout column). All matching selects are set, and the
visible one is read back to verify.

### Setting values on a React form

Assigning `el.value` does not update React state. Every write goes through:

```js
const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
setter.call(el, value);
el.dispatchEvent(new Event('input',  { bubbles: true }));
el.dispatchEvent(new Event('change', { bubbles: true }));
```

Checkboxes use `el.click()` only when `el.checked` differs from the target, so React's own handler runs.

### The seven step state machine

Each step declares the route it expects. If the current URL does not match, the content script does
nothing. That single guard prevents double firing on SPA re-renders.

| # | Step | Expects URL | Action |
|---|---|---|---|
| 1 | `NAV_TO_PRODUCTS` | `/dashboard` | click sidebar Products |
| 2 | `CLICK_CREATE` | `/products/` | click the top right Create, never the centered one |
| 3 | `FILL_PRODUCT` | `/products/add/` | Name = Servicios, Code = Código, Price = parsed price, Currency = USD (`value="2"`), Type of Item = Service (`value="2"`), Show In Website = No (`value="false"`), then Save |
| 4 | `OPEN_PRODUCT` | `/products/` | find the `li a` whose `h3` text equals the Servicios name and whose `p` contains the code and price, click it, store its UUID |
| 5 | `CLICK_CREATE_LINK` | `/products/<uuid>/permalink/product/` | verify `input[name=name]` and `input[name=code]` match the row, then click Create Fygaro Link for Product |
| 6 | `FILL_LINK` | `/payment-buttons/payments/payment-buttons/add/` | Name = Código, Currency = USD (`value="2"`), expand Advanced Options, check `require_phone`, `require_legal_id`, `require_billing_address`, then Save |
| 7 | `CAPTURE_LINK` | `.../payment-buttons/permalink/` | read the readonly textarea containing `https://www.fygaro.com/..../pb/...`, store it on the row, click Back to Home, advance the cursor |

Extra guards:
- Step 3 and 6 confirm the URL actually changed after Save. If it did not, the page rejected the form, so
  the run pauses with the on page validation text included in the error.
- Step 6 does not blindly click the Advanced Options toggle, because that button toggles and its label
  flips between `Show` and `Hide`. It checks for `input[name="require_phone"]` first and only clicks the
  toggle if the panel is not already open, then waits for the panel. Each of the three checkboxes is
  clicked only when `checked` is false, so re-running the step is safe.
- Step 5 is the reason the product UUID is captured in step 4: it proves we opened the right product
  rather than trusting recency.
- Step 7 reads the link from the textarea rather than using the Copy Link button, because clipboard reads
  need extra permissions and can silently fail. The Copy Link button is still clicked afterwards so the
  page behaves as the user expects.
- An action key of `step + url + cursor` prevents the same action firing twice within the debounce window.

### Random delay

Every click and every field write is preceded by `randomDelay(min, max)`, default 1200 to 3000 ms, both
adjustable in the UI. Step timeout defaults to 20 s and is also adjustable.

### Failure handling

On timeout, missing element, or a Save that did not navigate, the service worker sets
`status = 'awaiting_user'` and stores `pendingError { rowIndex, step, url, message, htmlHint }`. The side
panel shows a red banner with the row's Código and Servicios and three buttons: Retry step, Skip this row,
Stop run. A desktop notification fires so the user does not have to watch the screen.

## Design system (light theme, locked)

`:root { color-scheme: light; }` and no `prefers-color-scheme` rules anywhere, so the panel stays light
regardless of the OS or Chrome theme. Every value below is a CSS custom property defined once in
`sidepanel.css`, so the whole look is retunable from one block.

Palette, chosen to read as a finance and operations tool rather than a toy:

| Token | Value | Use |
|---|---|---|
| `--canvas` | `#F6F7F9` | panel background |
| `--surface` | `#FFFFFF` | cards, list rows, inputs |
| `--border` | `#E3E6EA` | hairlines, card edges |
| `--border-strong` | `#CBD2D9` | input borders, dividers |
| `--text` | `#1A1F26` | primary text |
| `--text-secondary` | `#5B6673` | labels, helper text |
| `--text-muted` | `#8A94A0` | timestamps, placeholders |
| `--primary` / `--primary-hover` / `--primary-tint` | `#2457D6` / `#1D46AD` / `#EEF3FF` | Start, primary buttons, active step, progress fill |
| `--success` / `--success-tint` | `#17803D` / `#E9F7EF` | done rows, completed run |
| `--warning` / `--warning-tint` | `#B45309` / `#FEF4E6` | paused, skipped rows |
| `--danger` / `--danger-tint` | `#C02626` / `#FDECEC` | failures, Stop, needs attention banner |

Type: `-apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif` for UI, and
`ui-monospace, "Cascadia Mono", Consolas, monospace` for Códigos, URLs and the log. Sizes are a small fixed
scale of 11, 12, 13, 15 and 18 px, weights 400, 500 and 600 only.

Structure: 4 px spacing base (4, 8, 12, 16, 24, 32), radius 6 px on controls and 10 px on cards, and a
single soft shadow `0 1px 2px rgba(16,24,40,.05), 0 1px 3px rgba(16,24,40,.06)`. Focus is always a visible
2 px `--primary` ring, never `outline: none`. Status is communicated with an icon and a word as well as a
colour, so it survives colour blindness and greyscale.

## Side panel UI

Single column, sectioned, keyboard accessible, `aria-live` on the status region.

1. **Header** - name, status pill (Idle / Running / Paused / Needs attention / Done), and whether a Fygaro
   tab is currently attached.
2. **Catalog** - drag and drop or click to pick the `.xlsx`. After parsing: sheet dropdown pre selected to
   `Logros ` (matched with trimming), column mapping shown as editable chips so a header rename does not
   break anything, and a summary line: total rows, to process, skipped because a Link already exists.
3. **Settings** - min and max delay, step timeout, and a **Dry run** toggle that performs every step
   including form fills but never clicks Save. This is how the user validates selectors safely.
4. **Run** - Start, Pause, Resume, Stop. A progress bar with `done / total`, a live ETA, and a card for the
   current row showing Código, Servicios, price, and a seven dot step breadcrumb with the active step lit.
5. **Results** - windowed list of all rows with a status badge (pending, done, skipped, failed) and the
   link, click to copy. Filter buttons by status.
6. **Export** - Download updated .xlsx (primary), Download CSV, Copy all links. Export is enabled as soon
   as one link exists, not only at the end.
7. **Log** - collapsible, timestamped, level coloured, capped at the last 500 entries, with Copy log.

The toolbar icon shows a badge with the completed count while running.

## The xlsx round trip, in vanilla JS

Reading:
1. Parse the ZIP central directory from the `ArrayBuffer`.
2. Inflate needed entries with the native `DecompressionStream('deflate-raw')`. Stored entries pass through.
3. `DOMParser` over `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/sharedStrings.xml` and the target
   sheet. Resolve `t="s"` through shared strings, `t="inlineStr"` through `<is><t>`, otherwise `<v>`.

Writing:
1. Take the original bytes, replace only the target sheet XML.
2. For each row that gained a link, set cell `H{row}` as `<c r="H.." t="inlineStr"><is><t>URL</t></is></c>`,
   inserted in correct column order, replacing the cell if it already exists. Using `inlineStr` means
   `sharedStrings.xml` is never touched, which removes a whole class of corruption risk.
3. Re-zip every entry with `CompressionStream('deflate-raw')` plus a small table driven CRC32, preserving
   entry names and order. No ZIP64 needed at this size.
4. Download as `Catálogo de Productos y Servicios Fygaro (con Links) YYYY-MM-DD.xlsx`. The original file on
   disk is never modified.

All other sheets, the table definition, drawings, styles and formatting survive untouched.

## Build order

1. Scaffold: `manifest.json`, icons, `README.md`, folder structure. Commit.
2. `price.js` plus `tests/price.test.mjs`, all five real formats green under `node --test`.
3. `zip.js`, `xlsx-read.js`, `xlsx-write.js` plus `tests/xlsx.test.html`, verified against the real 699 row
   workbook and re-opened in Excel.
4. `state.js` and the service worker: state schema, messaging, badge, notifications.
5. `content/dom.js` and `content/steps.js`, verified by `tests/selectors.test.html` against the seven saved
   DOM snapshots in `temp/html/` before any live run.
6. `content/main.js`: route watcher, tick loop, locks, resume.
7. Side panel HTML, CSS and JS.
8. End to end dry run, then a live run limited to the first row, then hardening.
9. README with full docs and developer details.

Commits happen at the end of each numbered phase, locally, never pushed.

## Verification

- `node --test tests/price.test.mjs` proves every real price string in the sheet parses correctly. I will
  assert against all 699 extracted values, not a sample.
- Open `tests/selectors.test.html` in Chrome. It injects each of the eight snapshots from `temp/html/` and
  asserts that every locator finds exactly the intended element, including that the centered Create button
  is rejected and the top right one is chosen, that the three `require_*` checkboxes resolve from
  `link_advance.txt`, and that the `max_successful_payments` checkbox is never confused with the number
  input of the same name.
- Open `tests/xlsx.test.html`, pick the real workbook, confirm 699 rows parse with correct Código,
  Servicios and price, patch a few H cells, download, and open the result in Excel to confirm nothing is
  corrupted and other sheets are intact.
- Load the unpacked extension, turn on **Dry run**, and run against the live logged in Fygaro app. Every
  step must be reached and every field must be filled correctly with zero records created.
- Turn off Dry run and run a single row end to end. Confirm the product exists with Service type, USD,
  Show In Website No, the link is created with the Código as its name and the three advanced fields
  checked, and the URL lands in the correct spreadsheet row.
- Then run the full set, with pause and resume exercised at least once mid run.

## Risks and how each is contained

| Risk | Containment |
|---|---|
| Fygaro ships a redeploy and hashed class names change | Every locator is driven by `name`, `href` or text. Hashed classes are only a last resort fallback. |
| A run is interrupted after 400 rows | State lives in `chrome.storage.local` and links are stored the moment they are read. Restart resumes at the cursor, and rows with links are skipped. |
| Something is created with wrong data | Dry run mode fills every field but never clicks Save, so the whole flow is provable before a single record exists. Step 5 verifies name and code on the product page before creating the link. |
| Patched xlsx opens corrupted | Links are written as `inlineStr`, never touching `sharedStrings.xml`. The original file is never modified, only a new download is produced. CSV export and copy all links are always available as a fallback. |
| Duplicate Códigos or Servicios names | Product lookup matches name plus code plus price and takes the first hit in a newest first list. |
| Run takes many hours | Pause and resume, persistent progress, a toolbar badge and a desktop notification when it finishes or needs attention. |
