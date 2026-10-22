## 9. Documents

How CodingCops reports, proposals, policies, letters and any A4 PDF or Word/Google doc look and read. Documents are calmer than slides and posts: they are read closely and often printed. So the brand shows up on the cover and in a few strong accents, and the inside stays quiet and legible. Start from the `A4Page` component.

### Page setup

- **Size:** A4 portrait, 210 × 297 mm. Use landscape only for wide tables or dashboards.
- **Margins:** 20 mm on every side, with 25 mm at the bottom to leave room for the footer.
- **Grid:** 6 columns with 5 mm gutters. Text runs across all 6 columns. Pull-quotes, figures and callouts can span 4 columns.
- **Export:** export PDFs as tagged PDF/A with fonts embedded and the document title set in the file properties. Name files `CodingCops_<Doc-type>_<Topic>_<YYYY-MM>.pdf`.

### Type (in points)

| Role | Style |
|---|---|
| Cover title | Satoshi Black, 40–48 pt, line height 1.0 |
| H1 (section) | Satoshi Bold, 24 pt, with a 12 pt `emerald` rule (3 pt thick) above it |
| H2 | Satoshi Bold, 15 pt |
| H3 | Satoshi Bold, 11.5 pt |
| Body | Satoshi Regular, 10.5 pt, line height 1.5, max 75 characters per line |
| Eyebrow / labels | Satoshi Medium, 8 pt, caps, letter-spacing +0.5 pt, `full-spectrum-blue` |
| Captions, footnotes | Satoshi Regular, 8.5 pt, `slate` |
| Script accent | Brittany Signature, 22–32 pt, **cover and closing page only** |

- Never set body text below 9 pt.
- Align text left; never justify it.
- Use one space after full stops.

### Colour

- Interior pages are **`white`**, not platinum, so they print cleanly. Text is `prussian-blue`, and secondary text is `slate`.
- Use `prussian-blue` or `full-spectrum-blue` for the cover, section-opener pages and one highlight box per spread.
- Use `emerald` for H1 rules, the one key figure on a page, and chart highlights. Never use it for text on white (2.2:1).
- `vibrant-coral` is reserved for warnings and "action required" labels, always paired with a word.
- The brand guide gives RGB values only. Ask the brand designer for official CMYK/Pantone values before commercial printing, rather than converting by eye.

### Cover page

- Use a full-bleed `prussian-blue` background, the **inline white logo** at top-left (40 mm wide), and an emerald blob bleeding off the bottom-right corner.
- Bottom-left, from top to bottom:
  - eyebrow (document type)
  - cover title (two lines at most)
  - a subtitle in `full-spectrum-blue` or `emerald`
  - the meta line: date, version, prepared by, classification ("Internal", "Confidential")
- Add one optional script line under the title.

### Interior pages

- **Header:** the monogram (6 mm), the document title in eyebrow style and the section name at the right, with a 0.5 pt `mist-border` rule under it.
- **Footer:** "CodingCops · Confidential" at the left and "Page n of N" at the right, in 8 pt `slate`.
- Do **not** repeat the full logo on interior pages.
- Start each H1 section on a new page in reports longer than 6 pages.

### Elements

- **Callout box:** `white` fill, a 1 pt `prussian-blue` outline and 4 mm corner radius, with an eyebrow label ("NOTE", "KEY TAKEAWAY"). A `prussian-blue` fill is allowed once per spread for the main takeaway.
- **Key figures:** at most 3 in a row, with the number in Satoshi Black 28 pt and the label in eyebrow style.
- **Tables:**
  - the header row in Satoshi Bold 9 pt on `blue-wash`, with a 1 pt `prussian-blue` rule under it
  - 0.5 pt `mist-border` row rules; no vertical lines or zebra stripes
  - numbers right-aligned
- **Charts:** follow the Presentations data rules. One `emerald` highlight, the rest `blue-wash` with `prussian-blue` outlines, direct labels, and a figure number plus source line in caption style.
- **Lists:** an emerald dash bullet, nested at most two levels. Keep a list to 7 items or fewer; past that, use a table.
- **Links:** `full-spectrum-blue`, underlined, and descriptive ("HR policy folder"). Never paste a raw URL in the body; put it in a footnote if the document will be printed.
- **Photos:** use the same treatment as Social, but smaller. One cutout or duotone image per section at most, with a caption.

### Letterhead and short documents

- **Page 1:** the inline blue logo top-left (35 mm), with the address, website and phone top-right in 8 pt `slate`, and a 3 pt `emerald` rule at the very bottom edge.
- **Page 2 onwards:** the monogram only, top-left.
- Order a letter's content as date, recipient block, subject in H3, body, then sign-off. A script signature is allowed in place of a scanned one only for celebratory letters.

### Word and Google Docs

- **Fonts:**
  - Install Satoshi for Word and embed fonts when saving (File → Options → Save → Embed fonts).
  - Google Docs can't load Satoshi. Use **Plus Jakarta Sans** as the closest available stand-in, and export to PDF from a Satoshi master for anything client-facing.
- **Templates:** keep one master template per type (Report, Proposal, Policy, Letter) with Heading 1–3, Normal, Caption and Table styles mapped to the table above. Never format by hand.

### Don't

- Don't use full-bleed colour on interior pages, background images behind text, or a logo on every page.
- Don't use more than two accent colours on one page, or emerald text on white.
- Don't use script outside the cover, or stretched or recoloured logos.
