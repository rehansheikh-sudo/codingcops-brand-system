# Admin setup: GitHub + Claude

One-time setup for whoever owns the brand system (brand admin + GitHub org
owner + Claude workspace owner). It takes about 30 minutes.

---

## 1. Put the repository on GitHub

1. Create a GitHub organization for CodingCops if you don't have one yet
   (github.com → your avatar → **Your organizations → New organization**).
2. Create a **private** repository in the organization called
   `codingcops-brand-system`, **without** a README (this repo already has one).
3. Push this folder:

   ```bash
   cd codingcops-brand-system
   git init -b main
   git add .
   git commit -m "CodingCops brand system v1.0"
   git remote add origin https://github.com/<YOUR-ORG>/codingcops-brand-system.git
   git push -u origin main
   ```

4. In `.github/CODEOWNERS`, replace `YOUR-ORG` with your organization name and
   push that change.

## 2. Permissions: employees can read, only brand admins can change

**Organization → Settings → Member privileges → Base permissions:** set to
**No permission**, so access comes only from teams.

**Organization → Teams**, create two teams:

| Team | Repo role on `codingcops-brand-system` | Who |
|---|---|---|
| `brand-admins` | **Maintain** (or Admin for the repo owner) | Noor + whoever approves brand changes |
| `all-employees` | **Read** | Everyone who makes branded material |

Add the teams in **Repository → Settings → Collaborators and teams**. Check
that nobody gets Write access through another team or as an outside
collaborator.

Read access lets people clone, pull and download releases. It does **not**
let them push. Note that read access still lets people copy the files; it
protects the official version, not secrecy.

## 3. Protect `main`

**Repository → Settings → Rules → Rulesets → New branch ruleset**

- Name: `Protect official brand` · Enforcement: **Active** · Target: default branch
- ✅ Restrict deletions · ✅ Block force pushes
- ✅ Require a pull request before merging: 1 approval, **require review from
  Code Owners**
- Bypass list: `brand-admins` (only if admins need to hot-fix without review)

> Rulesets on **private** repositories need GitHub Team or Enterprise. On the
> Free plan, skip this step: the Read role already stops employees from pushing.

## 4. Automatic skill releases

The workflow in `.github/workflows/release-skill.yml` runs on every merge to
`main` that changes the skill. It checks the skill and publishes
`codingcops-brand.zip` as a **Release**. Make sure Actions are allowed for the
repo (**Settings → Actions → General**). Run it once by hand from the
**Actions** tab (**Release brand skill → Run workflow**) to create the first
release.

## 5. Give everyone the skill in Claude.ai

Download `codingcops-brand.zip` from the latest release (or build it locally
with `./scripts/build-skill-zip.sh`).

**Team or Enterprise plan (recommended):** an organization owner opens
**Organization settings → Skills**, makes sure **Code execution and file
creation** and **Skills** are on, then clicks **+ Add** in *Organization
skills* and uploads the zip. Every member gets it right away. Members can turn
it off for themselves but can't change it.

**Pro or Max plans:** each person uploads the zip in **Settings →
Capabilities → Skills → Upload skill**. It's private to their account.

**After every brand update:** upload the new zip again so Claude.ai matches
`main`.

Docs: [Managing skills for your organization](https://support.claude.com/en/articles/13119606) ·
[Using skills in Claude](https://support.claude.com/en/articles/12512180-using-skills-in-claude)

## 6. Claude Code users

Nothing to do centrally. Each person follows the README: clone the repo, then
run `scripts/install-claude-code.sh` (or `.ps1` on Windows). After that,
`git pull` keeps them current.

## 7. Font licences: please confirm

The fonts are bundled so Claude can render and check its work, and so the
install scripts can put them on employees' machines.

- **Satoshi** is from Indian Type Foundry via Fontshare (ITF Free Font
  Licence). Free for commercial use; ITF asks to be credited as the owner of
  the typeface. Terms: <https://fontshare.com/terms>
- **Brittany Signature** is © Creatype Studio. Confirm that CodingCops holds
  a commercial licence that covers sharing the file with employees. Terms:
  <http://www.creatypestudio.co/license>

## 8. Test it (10 minutes)

- [ ] As a test employee with **Read** only: clone works, `git push` is rejected.
- [ ] Claude Code: run the install script, then in an empty folder ask
      `/codingcops-brand make a 6-slide pitch deck for a fintech prospect`.
- [ ] Claude.ai: with the skill on, ask "Use the CodingCops brand to make a
      one-page proposal for a fintech prospect".
- [ ] Open both outputs. Check the logo, colours, Satoshi, no shadows and the
      footer, then run `check_brand.py` on them.

## 9. Changing the brand later

1. Create a branch; edit `guidelines/*.md`. If a colour, font, size or logo
   changes, update `tokens/brand-tokens.json` and `tokens/brand.css` in the
   same change.
2. Run `python3 .claude/skills/codingcops-brand/scripts/cc_deck.py --demo /tmp/demo.pptx --preview /tmp/prev`
   and look at the previews.
3. Open a pull request. A brand admin approves, then merge.
4. Upload the new release zip in Claude.ai (step 5). Claude Code users pull.
