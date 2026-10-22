## 6. Components

Live previews and code are in the CodingCops design system. The summaries below are from each component's guidelines.

### Button

A pill-shaped action with a 1.5px `prussian-blue` outline. Use `primary` (`emerald` fill, `prussian-blue` text) **at most once per view**, for the main action on the page. Use `ghost` for secondary actions next to it, and `inverted` (`white`) for buttons on `full-spectrum-blue` or `prussian-blue` bands.

- Provide `variant`, a label that starts with a verb ("Start a project", "Book a call") and any native button props.
- On hover the button lifts 2px; on press it drops 2px. Never add a shadow.
- Never fill a button with `vibrant-coral` or `full-spectrum-blue`, and never put white text on emerald.

### LinkPill

A full-width link button for link lists and footers. It has a `white` (or `platinum`) fill, a 1.5px `prussian-blue` outline, `radius-pill`, and a label in `sans` 500.

- Provide `href`, a short label and an optional `icon`: a 20px monoline SVG stroked in `prussian-blue`.
- Stack link pills in one column up to 480px wide, with `space-12` gaps.

### Tag

A small pill label for categories and status.

- `neutral` (white) is for categories, `emerald` for live or active items, `coral` for new or highlighted ones, and `blue` (white text) for featured topics.
- Provide a `tone` and a one- or two-word label. Tags are not interactive.

### Card

A grid tile with a brand `fill`, a 2px `prussian-blue` outline and `radius-card` (or `radius-card-lg` with `size="lg"`). It has `space-24` padding and no shadow.

- Provide a `fill`, an optional `title` and children.
- Text colour follows the fill: `prussian-blue` on white, platinum, emerald and coral; `white` on blue and prussian.
- Arrange cards in asymmetric 2×3 or 3×3 grids, with some cards two columns wide. Use at most four fills in one grid.

### Input

A text field on `white` with a 1.5px `prussian-blue` outline and `radius-nav` (8px). Placeholder text uses `placeholder`.

- Provide native input props, plus an optional `label` and `hint` (in `caption` style, `slate`).
- On focus, the outline thickens to `stroke-bold` and a 2px `full-spectrum-blue` ring appears, offset by 2px.

### NavBar

A white bar that floats centred at the top of the page. It has a `white` fill, a 1.5px `prussian-blue` outline, `radius-nav` and `8px 8px 8px 24px` padding.

- Provide `logoSrc` (the URL of `codingcops-inline-blue.png`; it renders 24px tall), up to five `links` and one `action` (usually a primary Button).
- Float it over the hero with a `space-16` margin. Never stretch it into a full-width header.

### Banner

A full-width announcement strip, 44px tall, with an `emerald-wash` or `prussian-blue` fill and centred `body-sm` 500 text.

- Provide one short sentence and at most one underlined link.
- Place it above the nav. Use only one banner per page.

### SocialPost

A 4:5 feed-post template, designed at 1080 × 1350 and scaled to `width`. It lays out the house elements described in **Social**:
- a blue-washed background photo
- an emerald blob
- a cutout subject in a white sticker frame
- a coral paper label
- one script line
- a torn-paper strip
- a caps title with an emerald subline
- the monogram / script-word footer with an optional apply pill

- **Provide:**
  - the background `photo` and the `subject` cutout (a transparent PNG of the person)
  - `monogramSrc` (`codingcops-monogram.png`)
  - only the text slots the post needs
- Pick one of these per post, by series: a hiring post takes the `strip` and `cta`; a culture post takes the `title` and `subtitle`; a celebration takes the `label` and the `footerWord` + `footerTag` pair.
- At most one `script` line and one `strip` per post. The template is a layout guide for designers and quick mockups; final posts are finished in your design tool.

### Slide

A 16:9 deck template, designed at 1920 × 1080 and scaled to `width`. It follows the rules in **Presentations**. Pick a `variant`:

- `title`: the opening slide, with a dark ground, the inline white logo, a two-line headline, a script line, an emerald blob and an outlined period `mark` ("Q2").
- `section`: a divider with a rotating `tone` (blue, emerald, dark), a giant outlined number and one script line.
- `number`: one hero number with a takeaway headline.
- `content`: a takeaway headline, up to 4 `kpis` (one may be `highlight`) and one bar chart with a single highlighted bar and one `note`.
- `story`: a cutout `subject` over a blue-washed `photo`, with up to 3 `points`.
- `closing`: the same as `title`, for "Thank you".

- **Provide:**
  - the takeaway headline (a sentence, not a topic)
  - the asset URLs: `logoSrc` (inline white) and `monogramSrc`
  - only the slots the variant uses
- Content slides use the monogram footer and no logo. Keep detail in the speaker notes, not on the slide.

### A4Page

An A4 portrait page template, designed at 794 × 1123 px (A4 at 96 dpi) and scaled to `width`. It follows the rules in **Documents**.

- `cover`: a `prussian-blue` page with the inline white logo, an emerald blob, the title block and the meta line.
- `content`: an interior page with the monogram running header, the body and the "Confidential · Page n of N" footer.
- `letterhead`: the inline blue logo, the address block and a bottom emerald rule.

- **Provide:**
  - the logo URLs (`logoSrc`, `monogramSrc`)
  - the text slots
  - the body as `children`, using the page classes: `cc-a4-h1`, `cc-a4-h2`, `cc-a4-p`, `cc-a4-figs` (with `cc-a4-fig`, plus `hi` on one), `cc-a4-table` (`n` for number cells) and `cc-a4-callout`
- On screen this is a layout reference. Final PDFs come from the Word/InDesign master templates set up to these specs.
