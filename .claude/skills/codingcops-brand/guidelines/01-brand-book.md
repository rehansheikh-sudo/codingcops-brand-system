## 1. Brand book

CodingCops builds software with brilliance and precision, so the system pairs a confident tech palette with friendly, rounded shapes. Bold rounded cards in emerald, full-spectrum blue and coral sit on a platinum ground. Each card has a fine Prussian-blue outline and there are no shadows. There is one action colour.

### Logo

- Use the files in **Logos**. Never redraw, recolour, stretch or re-letter the mark or the wordmark.
- **Inline** (mark + wordmark side by side) is the default for headers, nav and documents. Use **Stack** in square or centred spaces: covers, social avatars with room, title slides.
- On `white`, `platinum` and the pale washes, use the **blue** versions. On `prussian-blue` or photography, use the **white** versions; their mark keeps its emerald and blue.
- Use the **Monogram** alone only where the name is already present or space is under 120px wide, such as favicons, app icons and small badges.
- Clear space is at least the height of the wordmark's letters on every side. Minimum inline width is 120px.

### Colour

- The ground is `platinum`. Cards, inputs and the nav sit on `white`. All text and outlines on light grounds are `prussian-blue`, and muted helper text is `slate`.
- `emerald` is the only action colour. Use one primary button per view, with `prussian-blue` text on it and never white.
- `full-spectrum-blue` is the brand band and feature colour. Use it for full-bleed sections, feature cards, links and the focus ring.
- `vibrant-coral` adds warmth. Use it for decorative cards, tags and highlights, never for buttons. Its main job is to give the cool palette some contrast.
- Rotate full-bleed section bands through `platinum`, `full-spectrum-blue`, `prussian-blue` and `emerald-wash`. The colour change acts as the divider.
- Text on fills:
  - `prussian-blue` on `emerald`, `vibrant-coral`, `white`, `platinum` and the washes
  - `white` on `full-spectrum-blue` and `prussian-blue`
  - `emerald` or `vibrant-coral` on `prussian-blue` for headline accents
- Never put white text on `emerald` (2.4:1). On `vibrant-coral`, white is allowed only at 24px and up.
- `logo-blue` is the blue inside the logo files and is for reference only. Build UI with `full-spectrum-blue`.

### Type

- One family, **Satoshi** (`sans`), for everything. Use the weights Light 300, Regular 400, Medium 500, Bold 700 and Black 900, each with an italic. There is no 600.
- Hero: `display` (Black, tight). Section titles: `heading` (Light), matching the airy titles in the brand guide. Card titles: `heading-sm` (Bold). Lead-ins: `subheading` (Medium).
- Text: `body` for running text, `body-lg` for hero copy, `label` (Bold) for buttons and `caption` (Medium) in `slate` for meta.
- **Accent: Brittany Signature** (`accent`), a handwritten script for warmth and design feel. Use it for about **3%** of the type on a page.
  - Set one or two words in `signature` or `signature-sm` beside or over a Satoshi headline ("Built *with care*"), or as a sign-off ("Thank you") or a quote flourish.
  - At most one accent per view, 32px or larger, in `prussian-blue`, `emerald` or `white` on dark bands.
  - Never use it for sentences, body copy, buttons, labels, navigation or data.
- Write headlines in sentence case. The wordmark is a separate typeface; never retype "CodingCops" in Satoshi as a logo. Use the files in Logos.

### Shape, outline and elevation

- Every card, button, input and nav has a `prussian-blue` outline: `stroke` (1.5px) on controls and `stroke-bold` (2px) on cards. Keep outlines fine: they define shapes, they shouldn't shout.
- **No shadows or blurs.** On hover an element lifts 2px, and on press it drops 2px.
- Radii:
  - `radius-pill` on everything interactive
  - `radius-card` or `radius-card-lg` on cards
  - `radius-blob` for discs, which echo the round monogram
  - `radius-nav` only on the nav and inputs
- Nothing gets a radius below 8px.

### Spacing and layout

- Content max-width is 1200px inside full-bleed bands. Bands pad `space-64` to `space-80` vertically.
- Card padding is `space-20` to `space-24`, and the element gap is `space-12` to `space-16`.
- In the hero, the headline goes left and the visual right. Later sections alternate sides. Card grids are asymmetric 2×3 or 3×3 clusters.

### States and focus

- Focus thickens the outline to `stroke-bold` and adds a 2px `full-spectrum-blue` ring offset by 2px (5.1:1 on platinum).
- Disabled controls get a `mist-border` outline and a `slate` label.

### Voice

- Write clear, confident and plain sentences. Address the reader as **you** and the company as **we**.
- Labels start with a verb: "Start a project", "Book a call", "See our work".
- Never use emoji as decoration.

### Iconography

- Use monoline icons with a 2px stroke in `prussian-blue`, 20–24px, with no fills. The brand guide doesn't set an icon library; use Lucide at stroke-width 2 until one is chosen.
