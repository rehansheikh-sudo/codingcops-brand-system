#!/usr/bin/env python3
"""
check_brand.py — flag off-brand colours, fonts, sizes and contrast in a .pptx or .docx.

    python3 check_brand.py deck.pptx
    python3 check_brand.py proposal.docx --tokens ../tokens/brand-tokens.json
    python3 check_brand.py --self-test

Checks (from tokens/brand-tokens.json):
  ERROR  a colour that is not in the CodingCops palette (or a reference-only logo colour)
  ERROR  a font other than Satoshi / Brittany Signature
  ERROR  a forbidden text-on-fill pair (e.g. white on emerald, emerald text on white/platinum)
  WARN   white/prussian text on coral/blue below the "large only" size
  WARN   text below the minimum size (9 pt), script accent below 32 px (16 pt in a deck)
  WARN   more than one script (Brittany Signature) text box on a slide
Standard library only. Exit code 1 when there are errors.
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
DEFAULT_TOKENS = HERE.parent / "tokens" / "brand-tokens.json"
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{%s}" % NS["w"]
ALLOWED_FONT_PREFIXES = ("Satoshi", "Brittany Signature", "+mj", "+mn")
SCRIPT = "Brittany Signature"


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, where, msg):
        self.items.append((level, where, msg))

    def print(self):
        errs = [i for i in self.items if i[0] == "ERROR"]
        warns = [i for i in self.items if i[0] == "WARN"]
        seen = set()
        for lvl, where, msg in self.items:
            key = (lvl, where, msg)
            if key in seen:
                continue
            seen.add(key)
            print(f"{lvl:5}  {where}: {msg}")
        print(f"\n{len(errs)} error(s), {len(warns)} warning(s)."
              + ("  On brand." if not errs and not warns else ""))
        return 1 if errs else 0


def load_tokens(path):
    t = json.loads(Path(path).read_text(encoding="utf-8"))
    palette = {v["hex"].lstrip("#").upper(): k for k, v in t["colors"].items()}
    reference = {v["hex"].lstrip("#").upper(): k for k, v in t.get("reference_only_colors", {}).items()}
    return t, palette, reference


def name_of(hexv, palette, reference):
    hexv = hexv.upper()
    return palette.get(hexv) or reference.get(hexv) or f"#{hexv}"


def check_colour(rep, where, hexv, palette, reference):
    hexv = hexv.upper()
    if hexv in ("AUTO",):
        return
    if hexv in reference:
        rep.add("ERROR", where, f"#{hexv} is {reference[hexv]} (reference only, inside logo files). Use the UI colour instead.")
    elif hexv not in palette:
        rep.add("ERROR", where, f"#{hexv} is not a CodingCops colour.")


def check_font(rep, where, face):
    if face and not face.startswith(ALLOWED_FONT_PREFIXES):
        rep.add("ERROR", where, f"font '{face}' is not Satoshi or Brittany Signature.")


def pair_rule(tokens, fill, text):
    rules = tokens["text_on_fill"].get(fill)
    if not rules:
        return None
    if text in rules.get("forbidden", []):
        return "forbidden"
    if text in rules.get("large_only", []):
        return "large_only"
    if text not in rules.get("allowed", []) and text not in rules.get("large_only", []):
        return "not_listed"
    return None


# ------------------------------------------------------------------ PPTX
def solid_hex(el):
    if el is None:
        return None
    c = el.find("a:solidFill/a:srgbClr", NS)
    return c.get("val").upper() if c is not None else None


def check_pptx(zf, tokens, palette, reference, rep):
    slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                    key=lambda n: int(re.findall(r"\d+", n)[-1]))
    for name in slides:
        where = "slide " + re.findall(r"\d+", name)[-1]
        root = ET.fromstring(zf.read(name))
        bg = root.find(".//p:bg/p:bgPr", NS)
        bg_hex = solid_hex(bg)
        if bg_hex:
            check_colour(rep, where + " background", bg_hex, palette, reference)
        for c in root.iter("{%s}srgbClr" % NS["a"]):
            check_colour(rep, where, c.get("val"), palette, reference)
        for lat in root.iter("{%s}latin" % NS["a"]):
            check_font(rep, where, lat.get("typeface"))
        script_boxes = 0
        filled = []  # (x0, y0, x1, y1, hex) of filled shapes, in z-order
        for sp in root.iter("{%s}sp" % NS["p"]):
            own = solid_hex(sp.find("p:spPr", NS))
            off, ext = sp.find("p:spPr/a:xfrm/a:off", NS), sp.find("p:spPr/a:xfrm/a:ext", NS)
            box = None
            if off is not None and ext is not None:
                x, y = int(off.get("x")), int(off.get("y"))
                box = (x, y, x + int(ext.get("cx")), y + int(ext.get("cy")))
            fill = own
            if not fill and box:  # text box with no fill: what is under its centre?
                cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
                for x0, y0, x1, y1, h in reversed(filled):
                    if x0 <= cx <= x1 and y0 <= cy <= y1:
                        fill = h
                        break
            fill = fill or bg_hex
            if own and box:
                filled.append((*box, own))
            fill_tok = palette.get(fill) if fill else None
            has_script = False
            for r in sp.iter("{%s}r" % NS["a"]):
                rpr = r.find("a:rPr", NS)
                txt = (r.findtext("a:t", default="", namespaces=NS) or "").strip()
                if rpr is None or not txt:
                    continue
                sz = int(rpr.get("sz", "1800"))
                face = (rpr.find("a:latin", NS).get("typeface") if rpr.find("a:latin", NS) is not None else "")
                if face.startswith(SCRIPT):
                    has_script = True
                    if sz < 1600:
                        rep.add("WARN", where, f"script accent '{txt[:30]}' is {sz/100:g} pt; minimum is 16 pt (32 px).")
                if sz < 900:
                    rep.add("WARN", where, f"text '{txt[:30]}' is {sz/100:g} pt; never below 9 pt.")
                th = solid_hex(rpr)
                if not (th and fill_tok):
                    continue
                text_tok = palette.get(th)
                rule = pair_rule(tokens, fill_tok, text_tok) if text_tok else None
                if rule == "forbidden":
                    rep.add("ERROR", where, f"{text_tok} text on {fill_tok} ('{txt[:30]}') is not allowed.")
                elif rule == "large_only" and sz < 1200:
                    rep.add("WARN", where, f"{text_tok} on {fill_tok} needs 24 px+ (12 pt); '{txt[:30]}' is {sz/100:g} pt.")
                elif rule == "not_listed":
                    rep.add("WARN", where, f"{text_tok} text on {fill_tok} ('{txt[:30]}') is not a listed pairing; check contrast.")
            script_boxes += has_script
        if script_boxes > 1:
            rep.add("WARN", where, f"{script_boxes} script text boxes; use one script accent per slide.")
    for name in [n for n in zf.namelist() if n.startswith("ppt/charts/chart") and n.endswith(".xml")]:
        root = ET.fromstring(zf.read(name))
        for c in root.iter("{%s}srgbClr" % NS["a"]):
            check_colour(rep, name.split("/")[-1], c.get("val"), palette, reference)
        for lat in root.iter("{%s}latin" % NS["a"]):
            check_font(rep, name.split("/")[-1], lat.get("typeface"))


# ------------------------------------------------------------------ DOCX
GLYPHS = {"–", "-", "—", "•", "·", "▲", "▼"}


def _fill_of(el):
    shd = el.find("w:shd", NS) if el is not None else None
    v = shd.get(W + "fill") if shd is not None else None
    return v.upper() if v and v.lower() != "auto" else None


def check_docx(zf, tokens, palette, reference, rep):
    names = zf.namelist()
    # theme fonts (used when styles say asciiTheme="minorHAnsi" etc.)
    theme_fonts = {}
    if "word/theme/theme1.xml" in names:
        th = ET.fromstring(zf.read("word/theme/theme1.xml"))
        for tag, key in (("majorFont", "major"), ("minorFont", "minor")):
            lat = th.find(f".//a:{tag}/a:latin", NS)
            if lat is not None:
                theme_fonts[key] = lat.get("typeface")
    # styles actually used (plus their basedOn chain), so unused built-ins are ignored
    content = [n for n in names if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)]
    used = {"Normal"}
    for n in content:
        r = ET.fromstring(zf.read(n))
        for tag in ("pStyle", "rStyle", "tblStyle"):
            used |= {e.get(W + "val") for e in r.iter(W + tag)}
    styles_root = ET.fromstring(zf.read("word/styles.xml")) if "word/styles.xml" in names else None
    style_els = {}
    if styles_root is not None:
        for st in styles_root.findall("w:style", NS):
            style_els[st.get(W + "styleId")] = st
            if st.get(W + "default") == "1" and st.get(W + "type") == "paragraph":
                used.add(st.get(W + "styleId"))
        frontier = list(used)
        while frontier:
            st = style_els.get(frontier.pop())
            if st is None:
                continue
            for tag in ("basedOn", "link"):
                b = st.find(f"w:{tag}", NS)
                if b is not None and b.get(W + "val") not in used:
                    used.add(b.get(W + "val")); frontier.append(b.get(W + "val"))

    def scan(where, el):
        for c in el.iter(W + "color"):
            v = c.get(W + "val")
            if v and v.lower() != "auto":
                check_colour(rep, where, v, palette, reference)
        for c in el.iter(W + "shd"):
            v = c.get(W + "fill")
            if v and v.lower() != "auto":
                check_colour(rep, where, v, palette, reference)
        for tag in ("pBdr", "tcBorders", "tblBorders"):
            for b in el.iter(W + tag):
                for side in b:
                    v = side.get(W + "color")
                    if v and v.lower() != "auto" and side.get(W + "val") not in ("nil", "none"):
                        check_colour(rep, where + " border", v, palette, reference)
        for rf in el.iter(W + "rFonts"):
            for attr in ("ascii", "hAnsi"):
                check_font(rep, where, rf.get(W + attr))
            for attr in ("asciiTheme", "hAnsiTheme"):
                t = rf.get(W + attr)
                if t:
                    face = theme_fonts.get("major" if t.startswith("major") else "minor")
                    check_font(rep, where + f" (theme font {t})", face)
        for c in el.iter("{%s}srgbClr" % NS["a"]):
            check_colour(rep, where, c.get("val"), palette, reference)

    if styles_root is not None:
        dd = styles_root.find("w:docDefaults", NS)
        if dd is not None:
            scan("styles.xml docDefaults", dd)
        for sid in sorted(x for x in used if x):
            if sid in style_els:
                scan(f"style '{sid}'", style_els[sid])

    for name in content:
        where = name.split("/")[-1]
        root = ET.fromstring(zf.read(name))
        scan(where, root)
        parent = {c: p for p in root.iter() for c in p}
        for r in root.iter(W + "r"):
            rpr = r.find("w:rPr", NS)
            txt = "".join(t.text or "" for t in r.iter(W + "t")).strip()
            if rpr is None or not txt:
                continue
            sz = rpr.find("w:sz", NS)
            if sz is not None and int(sz.get(W + "val")) < 16:
                rep.add("WARN", where, f"text '{txt[:30]}' is {int(sz.get(W + 'val'))/2:g} pt; labels and "
                        "captions stay at 8 pt+, body text at 9 pt+.")
            if txt in GLYPHS:
                continue
            col = rpr.find("w:color", NS)
            text_tok = palette.get((col.get(W + "val") or "").upper()) if col is not None else None
            fill = _fill_of(rpr)
            node = r
            while fill is None and node in parent:
                node = parent[node]
                if node.tag == W + "p":
                    fill = _fill_of(node.find("w:pPr", NS))
                elif node.tag == W + "tc":
                    fill = _fill_of(node.find("w:tcPr", NS))
            fill_tok = palette.get(fill) if fill else "white"
            if text_tok and fill_tok and pair_rule(tokens, fill_tok, text_tok) == "forbidden":
                rep.add("ERROR", where, f"{text_tok} text on {fill_tok} ('{txt[:30]}') is not allowed.")
            if text_tok == "white" and fill_tok == "white":
                rep.add("ERROR", where, f"white text on a white page ('{txt[:30]}').")


def check(path, tokens_path=DEFAULT_TOKENS):
    tokens, palette, reference = load_tokens(tokens_path)
    rep = Report()
    data = path if isinstance(path, (bytes, bytearray)) else Path(path).read_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = zf.namelist()
        if any(n.startswith("ppt/") for n in names):
            check_pptx(zf, tokens, palette, reference, rep)
        elif any(n.startswith("word/") for n in names):
            check_docx(zf, tokens, palette, reference, rep)
        else:
            rep.add("ERROR", str(path), "not a .pptx or .docx file")
    return rep


def self_test():
    slide = f"""<p:sld xmlns:p="{NS['p']}" xmlns:a="{NS['a']}"><p:cSld>
      <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F2F7F7"/></a:solidFill></p:bgPr></p:bg><p:spTree>
      <p:sp><p:spPr><a:solidFill><a:srgbClr val="01C16C"/></a:solidFill></p:spPr><p:txBody><a:p><a:r>
        <a:rPr sz="2000"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Book a call</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:spPr/><p:txBody><a:p><a:r>
        <a:rPr sz="1400"><a:solidFill><a:srgbClr val="01C16C"/></a:solidFill><a:latin typeface="Satoshi"/></a:rPr><a:t>green text</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:spPr><a:solidFill><a:srgbClr val="004CF8"/></a:solidFill></p:spPr></p:sp>
      <p:sp><p:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr></p:sp>
      </p:spTree></p:cSld></p:sld>"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("ppt/slides/slide1.xml", slide)
    rep = check(buf.getvalue())
    msgs = " | ".join(m for _, _, m in rep.items)
    expect = ["white text on emerald", "font 'Arial'", "emerald text on platinum", "logo-blue", "#123456"]
    missing = [e for e in expect if e not in msgs]
    if missing:
        print("SELF-TEST FAILED, not detected:", missing)
        print(msgs)
        return 1
    print(f"Self-test passed ({len(rep.items)} issues detected as expected).")
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", nargs="?")
    ap.add_argument("--tokens", default=str(DEFAULT_TOKENS))
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    if a.self_test:
        sys.exit(self_test())
    if not a.file:
        ap.error("give a .pptx or .docx file")
    sys.exit(check(a.file, a.tokens).print())


if __name__ == "__main__":
    main()
