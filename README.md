# Solar PV asset finance explainer

A single web page that turns a PV*SOL report into a five page A4 landscape explainer of 5 and 7 year asset finance, built to the Garland UK Brand Guidelines v2.0.

**How it is used**

1. Open the page.
2. Drop in the PV*SOL report (.docx or .pdf).
3. Type the lender's monthly repayment for the 5 year and 7 year options.
4. Click **Save as PDF**.

Everything else (client name, address, system details, savings, the 25 year cash flow, the chart and the sensitivity figures) is read from the report and calculated automatically. Anything the page could not find is flagged in orange so it can be typed in.

## Design basis

The layout follows the Garland UK Brand Guidelines v2.0 (2026).

- **Grid:** the 12 column A4 landscape grid, with 15 mm margins and a 5 mm gutter.
- **Type:** Roboto Black for headlines and large numerals only. Text is set in Helvetica Neue LT Pro when it is installed on the machine, with Arial as the approved substitute.
- **Colour:** Garland Blue and Black 6 for text, Garland Orange for section numbers and one key figure per page. The Solar PV gradient is used on the cover.
- **Devices:** the brand table style, the blue and light panels, and the blue hairline rule.
- **Logos:** the horizontal full colour logo on inner pages and the reversed white logo on the cover. Both were taken from the guidelines; swap in the master artwork files in `source/assets/` if available.

## Privacy

- The report is read inside the user's browser. It is never uploaded, saved or sent anywhere, and closing the tab clears it.
- The page has a security policy that blocks all outgoing connections, so it cannot send data even by accident.
- Every library, font and image is built into `index.html`. Nothing loads from outside.
- **Never commit real PV*SOL reports or finished explainers to this repository.**

## Publishing on GitHub Pages

1. Create a new repository (for example `garland-pv-finance`) and upload everything in this folder.
2. Go to **Settings > Pages**, set **Source** to *Deploy from a branch*, choose `main` and `/ (root)`, then save.
3. After a minute the page is live at `https://<your-account>.github.io/garland-pv-finance/`.
4. Optional: to use a Garland address such as `tools.garlanduk.com`, add it under **Custom domain** and ask IT to create the DNS record GitHub shows.

On a standard GitHub plan the page is public, although `robots.txt` and a `noindex` tag keep it out of search engines. It contains no client data, but if the blank template must not be public, host the same `index.html` behind a login instead (for example Cloudflare Pages with Cloudflare Access, or Azure Static Web Apps with Microsoft 365 sign in).

## Using it without hosting

`index.html` also works offline. Download it and double click to open it in Chrome or Edge.

## Making changes

The editable source is in `source/`:

| File | What it controls |
|---|---|
| `app.js` | Reading PV*SOL, the finance maths and all page wording |
| `styles.css` | Layout, colours and type (Garland palette) |
| `template.html` | The side panel |
| `assets/` | Logos and the cover photograph (replace `hero-solar.jpg` with a real Garland project photo when one is available) |

After editing, run `python3 source/build.py` to rebuild `index.html`, then commit both.

## What it reads from PV*SOL

Client name (company, project name or the first part of the address), site address, offer number, designer, report date, system size, module count and type, inverter count and model, roof areas, annual generation, PV used on site, self sufficiency, CO2 avoided, site consumption, investment, first year savings, energy price, energy price inflation, appraisal period and the yearly electricity savings from the cash flow table. If the cash flow table cannot be read, yearly savings are calculated from the first year saving and the inflation rate, and the explainer says so.

## Third party components

JSZip 3.10.1 (MIT), pdf.js 4.10.38 (Apache 2.0) and Roboto (Apache 2.0). Licences are in `source/vendor/`.
