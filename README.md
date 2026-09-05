# Fygaro Catalog Automation

A Chrome extension that turns a spreadsheet catalog into Fygaro products and payment links, then writes
every generated link back into the spreadsheet next to the row it came from.

The catalog behind this project holds 2073 services, most of them with a product photo. Creating each one
by hand means repeating a seven page workflow 2073 times. This extension performs that loop inside your own signed in browser session, with live
progress, pause and resume, and a patched spreadsheet at the end.

---

## Table of contents

- [What it does](#what-it-does)
- [Requirements](#requirements)
- [Installation](#installation)
- [Using it](#using-it)
- [Your spreadsheet](#your-spreadsheet)
- [Settings](#settings)
- [Products that already exist](#products-that-already-exist)
- [When something goes wrong](#when-something-goes-wrong)
- [Exporting](#exporting)
- [Starting over](#starting-over)
- [How it works](#how-it-works)
- [Project layout](#project-layout)
- [Testing](#testing)
- [Design](#design)
- [Privacy and safety](#privacy-and-safety)
- [Troubleshooting](#troubleshooting)
- [Developer](#developer)
- [License](#license)

---

## What it does

For every row of your catalog the extension performs these seven steps:

| # | Step | Page | What happens |
|---|------|------|--------------|
| 1 | Open Products | `/app/dashboard/` | Clicks Products in the sidebar |
| 2 | Start a new item | `/app/products/` | Clicks Create in the top right, never the centred one |
| 3 | Fill the product form | `/app/products/add/` | Name, Code and Price from the sheet, Currency USD, Type of Item Service, Show In Website No, then Save |
| 4 | Open the new product | `/app/products/` | Finds the product by name **and** code, never by assuming it is the newest |
| 5 | Start the Fygaro Link | `/app/products/<id>/permalink/product/` | Confirms the right product is open, then clicks Create Fygaro Link for Product |
| 6 | Fill the link form | `/app/payment-buttons/.../add/` | Name from the Código column, Currency USD, then ticks Phone, Legal ID and Billing Address under Advanced Options, then Save |
| 7 | Capture the link | `/app/payment-buttons/.../permalink/` | Reads the generated link, stores it against the row, returns to the dashboard |

Then it moves to the next row and starts again.

Every click and every keystroke is preceded by a random delay, so the automation never fires at a fixed
machine rhythm.

---

## Requirements

- Google Chrome 116 or newer, or any Chromium browser with side panel support
- A Fygaro account you are already signed in to
- Your catalog as an `.xlsx` file

There is nothing to install and nothing to build. The extension has **zero runtime dependencies**: it is
plain JavaScript, and the spreadsheet is read and written using the browser's own compression APIs.

---

## Installation

1. Open `chrome://extensions` in Chrome.
2. Turn on **Developer mode** using the switch in the top right.
3. Click **Load unpacked**.
4. Select this project folder, the one containing `manifest.json`.
5. Pin the extension so its icon stays visible, then click it to open the side panel.

---

## Using it

1. **Sign in to Fygaro** in a normal tab. The extension drives your session and never handles your
   password. Use **Open Fygaro** in the panel if you need a tab.
2. **Load your catalog.** Click the drop area in the panel, or drag the `.xlsx` onto it. The sheet named
   `Logros` and the Name, Code, Price and Link columns are detected automatically. Change any of them from
   the dropdowns if your file is laid out differently.
3. **Do a dry run first.** Tick **Dry run**, press Start, and the extension fills the product form for the
   first row and stops there without saving. Look at the form, confirm every field is right, then untick
   Dry run.
4. **Press Start.** Watch progress in the side panel, or from the small status widget that appears in the
   corner of the Fygaro page. You can carry on using other tabs, but leave the Fygaro tab open.
5. **Export** when you are ready. You do not have to wait for the whole run.

A run of 2073 rows takes many hours, so the Catalog card lets you set a row range and work through it in
batches. A run of that size takes tens of hours end to end. You do not have to sit through it: progress is written to disk after
every step, so you can pause, stop, close the panel, or even restart the browser and pick up where you
left off. Rows that already have a link are skipped, which is what makes resuming safe.

---

## Your spreadsheet

The extension reads these columns and writes back to one of them:

| Column | Header | Used as |
|--------|--------|---------|
| C | `Código` | The product **Code**, and the **Name** of the Fygaro Link |
| D | `Servicios` | The product **Name** |
| E | `Precio Total` | The product **Price** |
| H | `Columna 1` | The product **photo**, read but never written to |
| I | `Link` | Where the generated link is written. Created if it does not exist |
| J | `Nota` | Why a row finished without a link. Created if it does not exist |

Header names are matched ignoring case and accents, and each has accepted alternatives, so `Codigo`,
`Code`, `Precio` and `Price` all resolve. If a header is renamed beyond recognition you can map the columns
by hand in the panel.

**There is no Link column in the current catalog**, so the extension proposes one just past everything the
sheet uses, shows it in the Link dropdown labelled as new, and writes the `Link` header itself on export.
That header is one of the spellings it looks for, so the next run finds column I normally and skips the
rows that are already done. Column H holds the photos and is never written to.

**Photos.** Pictures anchored in column H are uploaded to the product's Gallery. A row without one is
created normally. Of the 2073 rows, 1030 carry a picture and 1043 do not, and although the whole pipeline
handles several per row, no row in this catalog has more than one. Only 35 distinct images are shared
across those 1030 rows, so each is held once rather than once per row.

**Prices.** The catalog mixes two number conventions in the same column and both are handled:

| In the sheet | Read as | Why |
|--------------|---------|-----|
| `B/.625,00` | 625.00 | The comma is the decimal separator |
| `B/.1.125,00` | 1125.00 | Dot groups thousands, comma is the decimal |
| `B/.100.00` | 100.00 | The dot is the decimal separator |
| `B/.1.000` | 1000.00 | A trailing group of exactly three digits reads as thousands |

When both a dot and a comma appear, whichever comes last is treated as the decimal point. Fygaro receives
a plain `1125.00` style value.

Rows missing a name, missing a code, carrying a price that cannot be read, or priced at zero are flagged
before the run starts and marked Failed, rather than breaking the run halfway through. Their reason is
written into the `Nota` column on export, so it is in the spreadsheet and not only in the panel.

Rows outside the chosen row range are set aside as Skipped, not Failed. Nothing is wrong with them; they
are simply not part of this batch.

---

## Settings

| Setting | Default | What it does |
|---------|---------|--------------|
| Min delay | 1200 ms | Shortest pause before a click or a keystroke |
| Max delay | 3000 ms | Longest pause. Each action picks a random time between the two |
| Step timeout | 20000 ms | How long to wait for a page before treating the step as failed |
| Attempts per step | 2 | How many times a step is retried before you are asked |
| Page zoom | 67 % | Zooms the Fygaro page out before the run starts, and puts it back afterwards |
| Dry run | off | Fills the first product form and stops without saving. Nothing is created |

**About the page zoom.** The side panel takes width away from the page. That can push Fygaro past a
responsive breakpoint and lay the forms out differently, which changes which copy of a duplicated control
is on screen. Zooming out gives that width back, so the page stays on the layout the steps were written
for. This is real browser zoom, the same thing Ctrl and minus does, because only real zoom changes the
layout viewport that the breakpoints respond to. Whatever zoom the page had before the run is remembered
and restored when the run stops or finishes. Set it to 100 to leave the page alone.

It is applied when you press Start, before the run navigates anywhere, and re-applied at the start of every
row so it cannot drift. The Activity log records it either way, including the reason if it could not be
applied.

---

## Products that already exist

If Fygaro refuses a Code because it is already in use, the row is skipped and the run carries on. It shows
as **Exists** in the Results list, and the export writes `Ya existe en Fygaro` into the `Nota` column. No
link is captured for it, and nothing about the existing product is changed.

Reloading the catalog later brings such a row back as pending, and it will be detected and skipped again.
That costs about a second rather than the forty seconds it used to take before this was recognised.

---

## When something goes wrong

The run pauses and asks you. It never guesses and never writes a link it is not sure about.

The side panel shows a red banner naming the row, the step and what actually happened, for example
`Saving the item was rejected by the page: Code already exists`. A desktop notification is raised too, so
you do not have to watch the screen. You then choose:

- **Retry this step** after fixing whatever was wrong on the page
- **Skip this row** to mark it Failed and carry on with the next one
- **Stop** to halt, keeping all progress

Everything captured so far stays saved, and you can export at any point.

---

## Exporting

Links go into column `I`, and the reason a row finished without one goes into column `J`. Both headers are
written on the first export if the sheet does not already have them. Column `H` and its photos are never
touched, and the table definition is left exactly as it is, so Excel opens the result without offering to
repair it.

You choose where the links go:

| Option | What happens |
|--------|--------------|
| **Save into a copy** (default) | Downloads a new `.xlsx`. Your original file is never touched. |
| **Update the original file** | Writes the links straight into the file you chose, in place. |

| Button | Produces |
|--------|----------|
| Save links to .xlsx | Either of the two above, depending on the option selected |
| Download CSV | Sheet row, Código, Servicios and Link, for pasting anywhere |
| Copy links | The same data on your clipboard, tab separated |

**Updating the original** needs the browser to grant write access to that exact file, which only the file
picker can do. Use the **Choose your catalog** button rather than dragging the file in, and the option
becomes available. Dragging still works, and Chrome hands over a handle for drops too where it can. If the
option stays greyed out, the text under it says why.

**Close the file in Excel first.** Windows will not let anything replace a file Excel is holding open, and
the panel will tell you so rather than failing quietly. The write is buffered and only committed at the
end, so an interrupted save cannot leave a half written spreadsheet behind.

Either way, everything else in the workbook survives: all other sheets, formatting, formulas, data
validations, drawings and the hyperlink that was already there. Only the cells in column H that gained a
link differ.

The workbook is cached inside the extension, so exporting still works after you close the panel or restart
the browser. If the cache is ever lost, load the same file again, or use the CSV export which needs no file
at all.

---

## Starting over

**Clear everything** in section 6 removes everything the extension has saved: the loaded catalog, every
captured link, the activity log, the cached copy of your file, its write permission, and your settings. The
panel tells you exactly what is stored before you press it, and asks for confirmation naming how many links
you are about to lose.

Your spreadsheet on disk is never touched by this. Export your links first if you still need them.

---

## How it works

Three parts, each with one job.

**The background worker** (`src/background/service-worker.js`) is the only thing that decides anything. It
owns the run and writes every decision to `chrome.storage.local` before acting, so a service worker that
Chrome shuts down mid run, or a browser restart, loses nothing.

**The content script** (`src/content/`) runs on Fygaro pages. It reports which page it is looking at, asks
the worker for a job, performs exactly that one step, and reports the outcome. It decides nothing itself.
It also draws the small status widget on the page, inside a shadow root so Fygaro's own styles and this
widget cannot affect each other.

**The side panel** (`src/sidepanel/`) is the interface. It loads the catalog, sends commands, and renders
whatever state comes back. Closing it cannot disturb a run. A side panel is used rather than a popup
precisely because a popup closes the moment you click away.

### Things that were designed for deliberately

- **Selectors survive redeploys.** Fygaro's class names are content hashed and change with every release,
  so no locator depends on one. Everything is found by `name`, by `href`, or by visible text, with text
  matching that folds case and accents so English and Spanish both work.
- **React is written to properly.** Assigning `element.value` is invisible to React. Every write goes
  through the native value setter and then dispatches real `input` and `change` events. Checkboxes are
  clicked rather than assigned.
- **The responsive duplicates are handled.** The product form renders Currency, Type of Item and Show In
  Website twice, once per layout column, with only one copy on screen. All copies are written and the
  visible one is read back to confirm.
- **A name collision is handled.** On the link form `max_successful_payments` is both a checkbox and a
  number input, so every input locator is qualified by type.
- **Hidden form controls are still found.** Fygaro paints its own checkbox and hides the real `input`
  behind it, so that input has no on screen box at all. Anything that reads or writes a form value
  therefore accepts a control that is not painted, while still preferring the on screen copy when a
  control is genuinely duplicated. Ticking is done by clicking the input, falling back to clicking its
  label, which is what a user actually does.
- **Advanced Options is not toggled blindly.** That button flips between Show and Hide. The extension
  checks whether the panel is already open instead of trusting the label.
- **Dropdown values are read, not hard coded.** USD, Service and No are found by their visible label, with
  the known value as a fallback, so Fygaro renumbering its options does not break anything.
- **Products are found by identity.** After saving, the product is located by matching both its name and
  its code, never by assuming the newest entry is the right one. The catalog contains repeated names, so
  this matters.
- **Every step confirms it worked.** A step that clicks Save waits for the page to actually move on. If it
  does not, the on page validation message is read and reported instead of a bare timeout.
- **The link is read, not pasted.** The generated URL is taken from the page itself rather than from the
  clipboard, which needs extra permissions and can fail silently. Copy Link is still pressed so the page
  ends in the state you would expect.
- **The zoom is applied before any work starts.** Chrome keys zoom to the origin, so a Fygaro page has to
  be loaded before the zoom can be set at all. The run therefore sets the zoom and then navigates, so the
  first page any step touches is already rendered at the intended layout. The original zoom is given back
  when the run ends, including when it ends by finishing rather than by being stopped.

---

## Project layout

```
manifest.json                        Manifest V3 definition
icons/                               Extension icons
src/shared/
  util.js                            Timing, text folding, CSV, formatting
  price.js                           Price parsing for both number conventions
  zip.js                             ZIP reader and writer over the native compression streams
  xlsx-read.js                       Workbook parsing: sheets, shared strings, rows
  xlsx-write.js                      Writes a column back without disturbing anything else
  state.js                           Run shape, storage keys, routes, message names
src/background/service-worker.js     Owns the run. Every decision is made here
src/content/
  dom.js                             Waiting, visibility, React aware writing
  steps.js                           Locators, and the seven steps
  main.js                            Route watching, the tick loop, the on page widget
src/sidepanel/                       The interface: html, css and js
tests/                               See below
tools/                               Test runner, syntax checks, fixture generator
docs/                                The catalog and the implementation plan
```

---

## Testing

```bash
npm test           # everything: node tests, then the browser pages
npm run test:node  # price parsing and the zip and xlsx layers
npm run lint       # syntax, house rules, and manifest references
```

`npm test` installs nothing. There are no dependencies to fetch.

| Suite | Checks | Covers |
|-------|--------|--------|
| `tests/price.test.mjs` | 9 | Every one of the 2073 real prices, cross checked against a separate reference implementation |
| `tests/state.test.mjs` | 11 | Run state, routes, and that a run saved by an older version gains every setting added since |
| `tests/xlsx.test.mjs` | 14 | ZIP round trips, style preservation, XML escaping, writing a column the sheet has no cells for, and that an unedited rewrite reproduces all 66 parts byte for byte |
| `tests/selectors.test.html` | 54 | Every locator, run against the nine captured page snapshots, with Fygaro's checkbox and file field styling reproduced |
| `tests/xlsx.test.html` | 23 | Workbook reading, sheet and column detection, the drawing that anchors the photos, patch and re read |
| `tests/integration.test.html` | 32 | The real side panel driving the real worker through a full run |

The integration suite is the interesting one. It stubs the Chrome APIs, loads the actual worker and the
actual panel, then plays a run through: 699 rows loaded, all seven steps walked for several rows, a failure
retried then escalated, a row skipped, pause and resume, a stalled navigation escalated, a reload prompt
declined without losing links, the page zoomed out before work begins and handed back on stop, zooming
retried after a failure, a blank settings field keeping its default rather than inventing one, a dry run,
a row range that sets the rest of the sheet aside, a picture served to the page in parts and rebuilt with a
matching checksum, thirty rows sharing one picture costing a single fetch, a duplicate code skipped without
a retry or a notification, an in place save checked cell by cell against a stand in file handle, a refused
write permission that
changes nothing, and a Clear everything that empties storage and puts the panel back to its first run
state.

The locator suite deliberately applies Fygaro's own checkbox styling to the snapshots, hiding the real
inputs behind painted replacements. Without that the snapshots render as plain visible checkboxes and a
locator that wrongly demands visibility passes the test while finding nothing on the live site. If you add
a snapshot, check whether the page styles that control before trusting a green run.

The browser suites run headless via `tools/run-browser-tests.mjs`. Set `CHROME_PATH` if Chrome is somewhere
unusual. The page snapshots come from `temp/html/` and are regenerated with `npm run fixtures`.

---

## Design

Light theme only. `color-scheme: light` is declared and there is no `prefers-color-scheme` rule anywhere,
so the panel looks the same whatever your operating system or Chrome theme is set to.

| Token | Value | Used for |
|-------|-------|----------|
| `--canvas` | `#F6F7F9` | Panel background |
| `--surface` | `#FFFFFF` | Cards, rows, inputs |
| `--border` / `--border-strong` | `#E3E6EA` / `#CBD2D9` | Hairlines and control edges |
| `--text` / `--text-secondary` / `--text-muted` | `#1A1F26` / `#5B6673` / `#8A94A0` | Type hierarchy |
| `--primary` | `#2457D6` | Primary actions, progress, the active step |
| `--success` / `--warning` / `--danger` | `#17803D` / `#B45309` / `#C02626` | Done, paused, failed |

Type is the system stack, with a monospace face for codes, links and the log. Sizes are a fixed scale of
11, 12, 13, 15 and 18 px and weights are limited to 400, 500 and 600. Spacing is a 4 px base, radii are
6 px on controls and 10 px on cards, and there is a single soft shadow.

Focus is always a visible ring, never removed. Status is conveyed by a word as well as a colour, so it
survives colour blindness and greyscale printing.

Every value is a custom property in one `:root` block in `src/sidepanel/sidepanel.css`, so the whole
appearance can be retuned from one place.

---

## Privacy and safety

- Your spreadsheet is read **in the browser** and is never uploaded anywhere.
- The extension talks to no server of its own. Its only host permission is `https://www.fygaro.com/*`.
- Your original file on disk is never modified. Exports are separate downloads.
- Nothing happens until you press Start. Dry run lets you prove the whole flow without creating a record.
- Your Fygaro credentials are never seen or handled. The extension uses the session you already opened.

---

## Troubleshooting

**The panel says a step could not be completed.** Read the banner: it names the row, the step and the
reason. Fix whatever it describes on the Fygaro page, then press Retry.

**"The Products link was not found in the sidebar."** The Fygaro tab is probably not on the dashboard, or
the session signed out. Sign in again, then press Retry.

**Nothing happens after pressing Start.** Check the Fygaro tab is open and signed in, and that the address
is under `https://www.fygaro.com/en/app/` or `/es/app/`. Use Open Fygaro in the panel.

**"Download updated .xlsx" is greyed out.** The cached workbook was lost. Load the same file again, or use
Download CSV, which needs no file.

**Excel will not save over the original.** Close the workbook in Excel first. The extension never writes to
your original file, so you are always replacing it yourself.

**A run seems stuck.** Raise Step timeout in Settings if Fygaro is responding slowly, then Retry.

**The page is not zooming, or zooms to the wrong level.** The zoom is applied when you press **Start**,
not when the panel opens, and the Activity log always says what happened: `Page zoom set to 67%`,
`Page zoom is already 67%, left as it is`, or a warning with the reason it failed. Check the Page zoom
value in Settings matches what you expect. Leaving that field blank keeps the value already in use rather
than guessing one.

**The Fygaro page is left zoomed out.** The zoom is restored when a run stops or finishes. If the browser
was closed mid run it can be left applied. Reset it with Ctrl and 0 on the Fygaro tab, or set Page zoom to
100 and start and stop a run.

**"The Advanced Options panel did not open."** The error now lists every checkbox actually on the form,
so compare that against `require_phone`, `require_legal_id` and `require_billing_address`. If the names
have changed, update `REQUIRED` in `src/content/steps.js`.

**After a Fygaro redesign.** Run `npm run test:browser`. The locator suite reports exactly which lookup
stopped matching.

---

## Developer

**Muhammad Abdullah Awais**
Full Stack Developer

- Website: <https://www.abdullahawais.com>
- Email: <contact@abdullahawais.com>
- LinkedIn: <https://www.linkedin.com/in/m-abdullah-awais-programmer>
- GitHub: <https://github.com/m-abdullah-awais>
- YouTube: <https://www.youtube.com/@m_abdullah_awais>
- Instagram: <https://www.instagram.com/m_abdullah_awais>

---

## License

Proprietary. All rights reserved.
