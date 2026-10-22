# CodingCops Brand System

The single, official source for the CodingCops brand: the design system,
logos, fonts, and the Claude skill that applies them.

Anyone at CodingCops can ask their own Claude to *"make a pitch deck in the
CodingCops brand"* and get the right logo, colours, Satoshi type, cards, slide
layouts and tone, without remembering the rules.

> **Read-only for employees.** Only the `brand-admins` team can change this
> repository. Generated work (decks, documents, posts) never goes in here.

---

## How it fits together

```
 GitHub: codingcops-brand-system  (private · employees Read · brand-admins approve changes)
        │
        ├──► Claude Code ── install script links the skill into ~/.claude/skills
        │                    and installs the fonts. `git pull` keeps it current.
        │
        └──► Claude.ai ──── each merge to main publishes codingcops-brand.zip (Releases)
                             → org owner uploads it once for everyone (Team/Enterprise)
                             → or each person uploads it (Pro/Max)
```

Everything Claude needs is in one self-contained skill folder:
[`.claude/skills/codingcops-brand/`](.claude/skills/codingcops-brand/). Claude
Code uses it directly, and the same folder is zipped for Claude.ai, so every
rule and asset exists in exactly one place.

---

## For employees

### Claude.ai (web, desktop app, Cowork)

1. Check **Settings → Capabilities**: *Code execution and file creation* must be on.
2. Look for `codingcops-brand` under **Skills**. If your admin added it for the
   organization, it's already there.
3. If it isn't, download `codingcops-brand.zip` from this repo's **Releases**
   and click **Upload skill**.
4. Ask for what you need:

   > Use the CodingCops brand to make a 10-slide pitch deck for a fintech
   > prospect. We cut their release time from monthly to weekly.

### Claude Code

```bash
git clone https://github.com/<YOUR-ORG>/codingcops-brand-system.git ~/codingcops-brand-system
~/codingcops-brand-system/scripts/install-claude-code.sh      # macOS / Linux
# Windows (PowerShell):  .\scripts\install-claude-code.ps1
```

The skill now works in **every** Claude Code session, and the brand fonts are
installed on your computer. Work in your own folder, not in this repo:

```bash
mkdir -p ~/work/acme-pitch && cd ~/work/acme-pitch
claude
> /codingcops-brand Make a 10-slide pitch deck for Acme. Save it here as acme-pitch.pptx.
```

Get brand updates: `git -C ~/codingcops-brand-system pull --ff-only`

### What you can ask for

| Ask for | You get |
|---|---|
| Pitch decks, presentations, QBRs | Editable .pptx built from the brand slide types (title, section, big number, KPI + chart, cards, story, closing) |
| Proposals, reports, policies, letters | A4 .docx / PDF to the document spec (cover, styles, header, footer, tables) |
| Instagram / LinkedIn posts, stories | 1080×1350 or 1080×1920 PNG layouts (add your own photos) |
| Landing pages, UI mockups, banners | HTML using the brand tokens and components |

Recipients of .pptx/.docx files need Satoshi installed to see them correctly.
Send a PDF when in doubt.

---

## For brand admins

- **Change a rule:** edit the section in
  `.claude/skills/codingcops-brand/guidelines/`, open a pull request, get it
  approved, merge.
- **Change a colour, font, size or logo:** update the guideline **and**
  `tokens/brand-tokens.json` **and** `tokens/brand.css` in the same pull request.
- **After merging:** a new `codingcops-brand.zip` appears under Releases.
  Upload it in Claude.ai (Organization settings → Skills).

First-time GitHub and Claude setup: **[docs/ADMIN-SETUP.md](docs/ADMIN-SETUP.md)**.

---

## Repository layout

```
codingcops-brand-system/
├── README.md                       ← you are here
├── CLAUDE.md                       ← rules for Claude when this repo itself is open in Claude Code
├── docs/ADMIN-SETUP.md             ← GitHub permissions, branch protection, Claude org skill
├── scripts/
│   ├── install-claude-code.sh      ← links the skill + installs fonts (macOS/Linux)
│   ├── install-claude-code.ps1     ← same for Windows
│   └── build-skill-zip.sh          ← builds dist/codingcops-brand.zip for Claude.ai
├── .github/
│   ├── CODEOWNERS                  ← brand-admins approve every change
│   └── workflows/release-skill.yml ← publishes the zip on every merge
└── .claude/skills/codingcops-brand/        ← THE SKILL (single source of truth)
    ├── SKILL.md                    ← how Claude applies the brand
    ├── guidelines/                 ← CodingCops Design System v1.0, one file per section
    ├── tokens/
    │   ├── brand-tokens.json       ← exact colours, contrast pairs, type scales, spacing, logos
    │   └── brand.css               ← the same as CSS variables + @font-face
    ├── assets/
    │   ├── logos/                  ← inline/stack (blue + white) and monogram PNGs
    │   ├── fonts/                  ← Satoshi (5 weights + italics), Brittany Signature
    │   └── shapes/                 ← the emerald blob (SVG + vector points)
    └── scripts/
        ├── cc_deck.py              ← builds on-brand .pptx decks from a simple spec
        ├── cc_docx.py              ← builds on-brand A4 .docx (cover, letterhead, styles, tables)
        └── check_brand.py          ← flags off-brand colours, fonts, sizes, contrast
```
