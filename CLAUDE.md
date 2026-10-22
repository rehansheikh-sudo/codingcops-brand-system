# CodingCops Brand System — instructions for Claude

This repository is the protected, official source of the CodingCops brand.

## If you are asked to create branded material

Use the `codingcops-brand` skill in `.claude/skills/codingcops-brand/`. Read its
`SKILL.md` first and follow it.

- Save every generated file **outside this repository** (the user's working
  folder, or `~/CodingCops-Brand-Output/` if none is given). Never save output
  inside this repo.
- Do not edit, rename, delete, commit or push anything in this repository while
  producing branded work.

## If a brand admin asks you to update the brand

Only do this when the user explicitly says they are updating the brand system.

- Keep the written rule and the machine-readable value in sync: a color or font
  change goes into `guidelines/*.md` **and** `tokens/brand-tokens.json` **and**
  `tokens/brand.css` in the same change.
- Every new asset or template must be listed in `SKILL.md` or the relevant
  guideline file, with when to use it.
- Never invent brand rules. If something is undecided, write it as an open
  question for the brand team instead of a rule.
- Work on a branch and open a pull request; `main` is protected.
- After the change, run `python3 .claude/skills/codingcops-brand/scripts/check_brand.py --self-test`
  and `scripts/build-skill-zip.sh` to confirm the skill still packages cleanly.
