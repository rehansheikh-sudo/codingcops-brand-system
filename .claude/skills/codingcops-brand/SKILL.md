---
name: codingcops-brand
description: Apply the official CodingCops design system (logos, colours, Satoshi and Brittany Signature type, cards and outlines, slide, document and social rules, tone of voice) to anything made for CodingCops. Use for any CodingCops-branded or company deliverable, including pitch decks, presentations, proposals, reports, letters, one-pagers, case studies, social posts, banners, posters, email headers, web pages or UI mockups.
---

# CodingCops Brand

Everything in this folder is approved by the CodingCops design team (Design
System v1.0, maintained by Noor). Follow it exactly. Where it is silent, pick
the quietest option that fits it and tell the user what you chose.

All paths below are relative to this skill's folder.

## 1. Load the brand first

1. Read `tokens/brand-tokens.json`: exact hex values, allowed text-on-fill
   pairs, type scales (web px, slide px, document pt), spacing, radii, logo files.
2. Read the guideline sections the job needs. They are the design system,
   split by section and unedited:

| Deliverable | Read |
|---|---|
| Anything | `guidelines/01-brand-book.md`, `guidelines/05-logos.md` |
| Colour or type decisions | `guidelines/02-colour.md`, `guidelines/03-typography.md` |
| Decks, presentations, pitch decks | `guidelines/08-presentations.md` |
| Proposals, reports, policies, letters, any A4/PDF/Word | `guidelines/09-documents.md` |
| Instagram / LinkedIn posts, stories, hiring posts | `guidelines/07-social.md` |
| Web pages, UI, components | `guidelines/04-spacing-radius-outlines.md`, `guidelines/06-components.md` |

## 2. Rules that are never broken

- **Logos:** place the real PNG from `assets/logos/`. Never redraw, recolour,
  stretch, re-letter or retype "CodingCops" as a logo. Blue versions on white,
  platinum and washes; white versions on prussian-blue or photos. Inline is the
  default, stack for square or centred spaces, monogram alone only where the
  name is already present or space is under 120px.
- **One action colour:** `emerald`. Text on emerald is always `prussian-blue`,
  never white. No emerald text on white or platinum.
- **Coral** (`vibrant-coral`) is decoration only: tags, highlights, small labels,
  one warm accent. Never a button fill.
- `logo-blue` (#004CF8) and the logo's emerald (#00C26C) are reference only.
  Use `full-spectrum-blue` and `emerald`.
- **Type:** Satoshi for everything (300/400/500/700/900; no 600). Brittany
  Signature is the accent only: about 3% of the type, one or two words, one per
  view, 32px or larger. Never for sentences, buttons, labels or data. Sentence
  case headlines.
- **Shapes:** prussian-blue outlines (1.5px controls, 2px cards), radius 8px or
  more, **no shadows, no blurs, no gradients on cards**.
- **Words:** clear, confident and plain. "You" for the reader, "we" for
  CodingCops. Labels start with a verb. No decorative emoji. Never invent client
  names, numbers, quotes or awards. Use visible placeholders like
  `[Client name]` or `[XX%]` for anything the user did not give you.
- **Output location:** save deliverables in the user's working or outputs
  folder, never inside the brand-system repository.

## 3. Build the deliverable

### Presentations and pitch decks → editable .pptx

Use the deck builder. It implements every slide type in `08-presentations.md`
(title, section, big number, bullets, KPI + chart, cards, story, closing), with
the brand theme, logos, emerald blob, footers and speaker notes.

```bash
python3 scripts/cc_deck.py --spec deck.json --out "<output folder>/Deck.pptx" --preview "<scratch>/previews"
```

`deck.json` is `{"slides": [{"type": "...", ...}]}`. Run
`python3 scripts/cc_deck.py --demo demo.pptx` once to see every type, or read
the `DEMO` spec and method docstrings at the bottom of `scripts/cc_deck.py`.
From Python: `sys.path.insert(0, "<skill>/scripts"); from cc_deck import Deck`.

Writing rules the builder can't enforce:
- Every headline is the **takeaway**, not the topic ("Defects fell 62% in six
  months", not "Quality results"). Two lines at most.
- At most 3 bullets per slide, about 12 words each. Detail goes in `notes`.
- One fact per chart, with the message bar highlighted (`highlight` index) and
  one `note`. Up to 4 KPI tiles; one tile may be the dark `highlight`.
- Use big-number slides for the deck's 3–5 most important facts.
- Sections rotate blue, emerald and dark automatically. Use one script line
  per title or section slide only.
- Pitch deck default flow: title → problem (section, big number) → solution
  (bullets/cards) → proof (KPI + chart, story) → how we work (cards) →
  closing. Shorten or adapt it to the request.
- The closing contact line must come from the user. Otherwise use
  `[contact email] · codingcops.com`. Never invent an email address or a
  person's name.
- Card fills: leave them at the defaults, or use at most four. Coral only as
  the one warm accent on a small card.

Story slides need a real photo with a duotone treatment. If the user has none,
leave the labeled placeholder in and say so.

Previews come from LibreOffice and are close to PowerPoint, not identical.
Outlined numbers are images, so they match exactly. Check the PNGs for
overflow and awkward wraps, then shorten the copy rather than shrinking type.

If a request needs a layout the builder doesn't have, build it with python-pptx
on the same canvas: 1920 × 1080 design px, **1px = 0.5pt = 6350 EMU**, 96px
margins, a 12-column grid with 24px gutters, and the footer on content slides
(monogram, section name and page number, with no logo).

### Documents (proposals, reports, policies, letters, one-pagers) → .docx, plus PDF if asked

Use the document builder. It sets up A4 pages, brand styles (Satoshi, with
Word's default Calibri and Office-blue removed), the brand theme, cover,
letterhead, header and footer with live page numbers, and the table, callout,
key-figure and list formats from `09-documents.md`:

```python
import sys; sys.path.insert(0, "<skill>/scripts")
from cc_docx import Doc, render_previews
d = Doc("PayLane payments app", doc_type="Proposal")       # kind="report" (default)
d.cover(subtitle="...", meta="October 2026 · v1.0 · Prepared by [Name] · Confidential")
d.h1("Summary"); d.p("..."); d.key_figures([("40+", "Fintech products"), ...])
d.h2("Plan"); d.bullets([...]); d.table(header, rows, numeric_cols=[1, 2], caption="Figure 1. ...")
d.callout("...", label="Key takeaway", dark=True)
d.save("<output folder>/CodingCops_Proposal_PayLane_2026-10.docx")
render_previews(path, "<scratch>/previews")                # PDF + PNGs to look at
```

- **Which layout:** use `cover()` for reports and proposals longer than about
  two pages. For letters, one-pagers and short documents, use
  `Doc(..., kind="letter", address=[...])` with `title_block()`. That gives the
  letterhead: logo top-left, address top-right, and an emerald rule along the
  bottom edge. Ask for the real address, or keep the placeholders.
- **Emerald on a page** goes on the H1 rule, the one key figure (as its accent
  rule, which `key_figures` adds) and chart highlights. Never as text on white.
- Callouts are square in Word because tables can't have rounded corners. Use
  the 4mm radius only in HTML-to-PDF output.
- Start each H1 on a new page in reports longer than 6 pages: `h1(text, new_page=True)`.
- Name PDFs `CodingCops_<Doc-type>_<Topic>_<YYYY-MM>.pdf`.

### Social posts → PNG at 1080 × 1350 (feed) or 1080 × 1920 (story)

Build as HTML with `tokens/brand.css` (it loads the brand fonts from
`assets/fonts/`) and screenshot it at exact size with a headless browser. Follow
`07-social.md`: blue-washed photo ground, a cutout person with a white sticker
outline, one emerald blob, caps Satoshi headline in white, and the monogram /
script-word footer. Real photos and cutouts must come from the user. Without
them, produce a layout mockup with labeled photo areas. Final posts are finished
in a design tool.

### Web pages, landing pages and UI mockups

Use `tokens/brand.css` variables and classes for every colour, font, radius and
outline. Follow `06-components.md` (pill buttons, floating nav, outlined cards,
asymmetric 2×3 or 3×3 card grids, bands rotating platinum → blue → prussian →
emerald-wash, content max-width 1200px). Icons: Lucide at stroke-width 2 in
prussian-blue.

### Anything else (posters, banners, email headers, event material)

Take the closest spec (a document cover for print, a social post for screens)
and apply the brand book. Say which spec you based it on.

## 4. Fonts in files you hand over

Satoshi and Brittany Signature are in `assets/fonts/`. Install them before
rendering previews (`cc_deck.install_fonts()` does this). PowerPoint and Word
files reference the fonts by name, so tell the user that recipients need
Satoshi installed, or that a PDF should go out instead. For Google Docs, Plus
Jakarta Sans is the approved stand-in.

## 5. Check before delivering

1. `python3 scripts/check_brand.py <file.pptx|file.docx>`. Fix every ERROR;
   fix or justify every WARN.
2. Render to images and look at every slide or page: logo placement and clear
   space, text overflow or awkward wraps, alignment, one accent per view, no
   empty or crowded areas.
3. Re-read the copy against the voice rules and remove anything invented.

## 6. Report back (two or three lines)

- What you made and where it is saved.
- Placeholders the user still needs to fill (`[Client name]`, photos).
- Anything the guidelines didn't cover and what you chose. Never call the work
  on-brand without having run the checks.
