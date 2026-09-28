#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXHIBITION LABEL STUDIO
=======================
Turns a Microsoft Forms / Excel / CSV export of artwork entries into printable
gallery labels (PDF).  Everything lives in this one file.

    python Generate_Labels.py               make the labels
    python Generate_Labels.py configure     open the friendly set-up screen

Folders (created automatically):
    EXCEL_FILES_HERE/   drop your .xlsx / .csv exports here
    profiles/           one small file per form layout  (e.g. "What Orana Means To Me - 2026.json")
    Generated_Labels/   finished PDFs appear here
"""
import sys, os, re, csv, copy, json, shutil, argparse, importlib, subprocess, textwrap, unicodedata, atexit
from xml.sax.saxutils import escape as xml_escape
from datetime import datetime


# ════════════════════════════════════════════════════════════════════════════
#  First-run helper: install the two libraries we need if they are missing
# ════════════════════════════════════════════════════════════════════════════
def _ensure_packages():
    needed = []
    for mod, pkg in (("openpyxl", "openpyxl"), ("reportlab", "reportlab")):
        try:
            importlib.import_module(mod)
        except ImportError:
            needed.append(pkg)
    if not needed:
        return
    print("\n  First-time setup: installing %s (needs internet, one time only)...\n" % ", ".join(needed))
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + needed)
    except Exception:
        print("\n  Could not install the required packages automatically.")
        print("  Please run:  pip install " + " ".join(needed))
        sys.exit(1)
    importlib.invalidate_caches()


_ensure_packages()

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

IS_WIN = os.name == "nt"
if IS_WIN:
    import msvcrt
else:
    import termios, tty, select

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ════════════════════════════════════════════════════════════════════════════
#  Paths & settings
# ════════════════════════════════════════════════════════════════════════════
BASE = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(BASE, "EXCEL_FILES_HERE")
OUTPUT_FOLDER = os.path.join(BASE, "Generated_Labels")
PROFILES_DIR = os.path.join(BASE, "profiles")
DATA_EXTS = (".xlsx", ".xlsm", ".csv")

MIN_FONT = 10.0          # statement text is never smaller than this (points)
PAD = 14                 # space inside each label edge (points)
RULE_GAP = 16            # space above the artist statement (points)
PAGE_SIZES = {"A4": A4, "A3": A3, "Letter": letter}

DEFAULT_LAYOUT = dict(page_size="A4", orientation="landscape", columns=2, rows=2, margin_mm=10,
                      statement_size=12, statement_min=10, sort="file")
DEFAULT_AUCTION = dict(yes_words=["yes", "y", "true"], text_yes="Available for Auction", text_no="")

# (key, short name, question, required, words that hint at the right column)
FIELDS = [
    ("artist", "Artist name", "Which column holds the ARTIST'S NAME?", True,
     ["full name", "artist", "participant", "student", "name"]),
    ("role", "Year level / class", "Which column holds the artist's YEAR LEVEL or CLASS?  (shown under the name)", False,
     ["year level", "year", "class", "grade", "role", "form"]),
    ("title", "Artwork title", "Which column holds the ARTWORK TITLE?", True,
     ["title", "name of your artwork", "artwork"]),
    ("statement", "Artist statement", "Which column holds the ARTIST STATEMENT?  (the longer paragraph)", False,
     ["statement", "describe", "about your", "story"]),
    ("auction", "Auction status", "Which column says whether the artwork is for the SILENT AUCTION?  (a Yes / No answer)", False,
     ["auction", "for sale", "sell"]),
    ("permission", "Display permission", "Which column holds PERMISSION TO DISPLAY?  (optional - you'll get a warning if anyone didn't say Yes)", False, ["display", "permission", "consent"]),
]
FIELD_NAME = {k: n for k, n, *_ in FIELDS}
FIELD_KEYS = tuple(k for k, *_ in FIELDS)

DEFAULT_PROFILE = {
    "profile_name": "What Orana Means To Me - 2026",
    "fields": {
        "artist": "Your full name (participant)",
        "role": "Your year level / class",
        "title": "Title of your artwork",
        "statement": "Artist Statement",
        "auction": "included in the silent auction",
        "permission": "I understand my artwork and artist statement may be displayed",
    },
    "details": [
        {"label": "", "column": "What mediums did you use?"},
        {"label": "", "column": "What are the dimensions of your physical piece?"},
    ],
    "auction": copy.deepcopy(DEFAULT_AUCTION),
    "layout": copy.deepcopy(DEFAULT_LAYOUT),
    "form_columns": [],
}

NA_VALUES = {"na", "n/a", "n.a.", "none", "nil", "-", "--", "not applicable"}


# ════════════════════════════════════════════════════════════════════════════
#  Small text helpers
# ════════════════════════════════════════════════════════════════════════════
def norm(s):
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip().casefold()


def header_key(h):
    """A short, stable name for a column header (its first line, lower-case)."""
    lines = [l.strip() for l in str(h).replace("\xa0", " ").splitlines() if l.strip()]
    k = norm(lines[0]) if lines else ""
    k = re.sub(r"(\.\.\.|…)+$", "", k).strip()
    return k[:80].strip()


def header_phrase(h):
    """What we save in a profile for a column: its first line, tidied, cut at a whole word (max 80 characters)."""
    lines = [l.strip() for l in str(h).replace("\xa0", " ").splitlines() if l.strip()]
    t = re.sub(r"\s+", " ", lines[0] if lines else "").strip()
    t = re.sub(r"(\.\.\.|…)+$", "", t).strip()
    if len(t) > 80:
        t = t[:80].rsplit(" ", 1)[0]
    return t


def short(s, n):
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s if len(s) <= n else s[: max(1, n - 1)].rstrip() + "…"


def nice_header(h, n=60):
    lines = [l.strip() for l in str(h).replace("\xa0", " ").splitlines() if l.strip()]
    return short(lines[0] if lines else "(blank column)", n)


def clean_text(val, multi=True):
    """Tidy one spreadsheet cell.  MS Forms ends 'tick all that apply' answers with ';' - turn those into commas."""
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    s = str(val).replace("\xa0", " ").replace("\u200b", "").strip()
    if multi and s.endswith(";"):
        s = ", ".join(p.strip() for p in s.split(";") if p.strip())
    return s


def is_yes(text, words):
    t = norm(text)
    return any(re.match(re.escape(norm(w)) + r"\b", t) for w in words if norm(w))


def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]+', "", str(name)).strip().rstrip(".") or "Untitled"


def stmt_markup(text):
    lines = [l.strip() for l in str(text).replace("\r", "").split("\n")]
    t = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    return xml_escape(t).replace("\n", "<br/>")


# ════════════════════════════════════════════════════════════════════════════
#  Reading spreadsheets
# ════════════════════════════════════════════════════════════════════════════
def find_data_files():
    if not os.path.isdir(INPUT_FOLDER):
        return []
    return [os.path.join(INPUT_FOLDER, f) for f in sorted(os.listdir(INPUT_FOLDER), key=str.casefold)
            if f.lower().endswith(DATA_EXTS) and not f.startswith(("~$", "."))]


def read_table(path):
    """Return (headers, rows).  rows = [(sheet_row_number, [cell, ...]), ...]"""
    if path.lower().endswith(".csv"):
        data, last = None, None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(path, newline="", encoding=enc) as fh:
                    data = list(csv.reader(fh))
                break
            except UnicodeDecodeError as e:
                last = e
        if data is None:
            raise last
    else:
        wb = load_workbook(path, data_only=True)   # (read-only mode mis-reads MS Forms files)
        try:
            data = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        finally:
            wb.close()
    numbered = [(i + 1, r) for i, r in enumerate(data)]
    while numbered and not any(clean_text(v) for v in numbered[0][1]):
        numbered.pop(0)
    if not numbered:
        return [], []
    headers = ["" if h is None else str(h) for h in numbered[0][1]]
    rows = [(n, r) for n, r in numbered[1:] if any(clean_text(v) for v in r)]
    return headers, rows


def find_col(headers, phrase):
    """Find the column matching a saved phrase: exact > starts-with > contained-in-header."""
    p = norm(phrase)
    if not p:
        return None
    keys = [header_key(h) for h in headers]
    for i, k in enumerate(keys):
        if k == p:
            return i
    for i, k in enumerate(keys):
        if k.startswith(p):
            return i
    for i, h in enumerate(headers):
        if p in norm(h):
            return i
    return None


def guess_column(headers, hints, taken=()):
    """Best-guess column for a field, using words that usually appear in its question."""
    best, best_score = None, 0
    for i, h in enumerate(headers):
        k, full = header_key(h), norm(h)
        if i in taken or not k or k in ("id", "email", "start time", "completion time", "last modified time"):
            continue
        for rank, w in enumerate(hints):
            score = (len(hints) - rank) if w in k else (len(hints) - rank) * 0.5 if w in full else 0
            if score > best_score:
                best, best_score = i, score
    return best


# ════════════════════════════════════════════════════════════════════════════
#  Profiles  (one JSON file per form layout, kept in ./profiles)
# ════════════════════════════════════════════════════════════════════════════
def normalize_profile(p):
    p.setdefault("profile_name", "Untitled")
    p["fields"] = {**{k: "" for k in FIELD_KEYS}, **p.get("fields", {})}
    p.setdefault("details", [])
    p["auction"] = {**DEFAULT_AUCTION, **p.get("auction", {})}
    p["layout"] = {**DEFAULT_LAYOUT, **p.get("layout", {})}
    p.setdefault("form_columns", [])
    return p


def new_profile(name):
    return normalize_profile({"profile_name": name})


def profile_path(name):
    return os.path.join(PROFILES_DIR, safe_filename(name) + ".json")


def save_profile(p):
    os.makedirs(PROFILES_DIR, exist_ok=True)
    path = profile_path(p["profile_name"])
    old = p.get("_file")
    data = {k: v for k, v in p.items() if not k.startswith("_")}
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    if old and os.path.abspath(old) != os.path.abspath(path) and os.path.exists(old):
        os.remove(old)
    p["_file"] = path


BAD_PROFILES = []


def load_profiles():
    """Load every profile.  Creates the folder (with the Orana example) the first time."""
    fresh = not os.path.isdir(PROFILES_DIR)
    os.makedirs(PROFILES_DIR, exist_ok=True)
    if fresh:
        save_profile(normalize_profile(copy.deepcopy(DEFAULT_PROFILE)))
    out, BAD_PROFILES[:] = [], []
    for fn in sorted(os.listdir(PROFILES_DIR), key=str.casefold):
        if not fn.lower().endswith(".json"):
            continue
        path = os.path.join(PROFILES_DIR, fn)
        try:
            with open(path, encoding="utf-8") as fh:
                p = normalize_profile(json.load(fh))
            p["_file"] = path
            out.append(p)
        except Exception:
            BAD_PROFILES.append(fn)
    return out


def resolve(profile, headers):
    idx = {k: find_col(headers, profile["fields"].get(k, "")) for k in FIELD_KEYS}
    det = [(d.get("label", ""), find_col(headers, d.get("column", "")), d.get("column", "")) for d in profile["details"]]
    return idx, det


def match_score(profile, headers):
    """0..1 - how many of this profile's columns exist in these headers."""
    want = [v for v in profile["fields"].values() if v] + [d["column"] for d in profile["details"]]
    if not want:
        return 0.0
    return sum(find_col(headers, w) is not None for w in want) / len(want)


# ════════════════════════════════════════════════════════════════════════════
#  Fonts  (uses a real Unicode font when one is installed, so accents & symbols work)
# ════════════════════════════════════════════════════════════════════════════
FONTS = {"reg": "Helvetica", "bold": "Helvetica-Bold", "ital": "Helvetica-Oblique", "cmap": None, "loaded": False}
_FAMILIES = [
    ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    ("liberationsans-regular.ttf", "liberationsans-bold.ttf", "liberationsans-italic.ttf", "liberationsans-bolditalic.ttf"),
    ("notosans-regular.ttf", "notosans-bold.ttf", "notosans-italic.ttf", "notosans-bolditalic.ttf"),
    ("dejavusans.ttf", "dejavusans-bold.ttf", "dejavusans-oblique.ttf", "dejavusans-boldoblique.ttf"),
    ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
]
_FONT_DIRS = [os.environ.get("WINDIR", r"C:\Windows") + r"\Fonts", "/usr/share/fonts", "/usr/local/share/fonts",
              os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts"),
              "/Library/Fonts", "/System/Library/Fonts/Supplemental"]


def load_fonts():
    if FONTS["loaded"]:
        return
    FONTS["loaded"] = True
    found = {}
    for d in _FONT_DIRS:
        for root, _, files in os.walk(d):
            for f in files:
                if f.lower().endswith(".ttf"):
                    found.setdefault(f.lower(), os.path.join(root, f))
    for fam in _FAMILIES:
        if all(f in found for f in fam):
            try:
                for name, f in zip(("LabelReg", "LabelBold", "LabelItal", "LabelBoldItal"), fam):
                    pdfmetrics.registerFont(TTFont(name, found[f]))
                FONTS.update(reg="LabelReg", bold="LabelBold", ital="LabelItal",
                             cmap=pdfmetrics.getFont("LabelReg").face.charToGlyph)
                return
            except Exception:
                continue


def unsupported(text):
    """Characters the label font cannot draw."""
    bad = []
    for ch in str(text):
        if ch in bad or unicodedata.category(ch) in ("Zs", "Cf", "Cc", "Zl", "Zp"):
            continue
        if FONTS["cmap"] is not None:
            ok = ord(ch) in FONTS["cmap"]
        else:
            try:
                ch.encode("cp1252")
                ok = True
            except UnicodeEncodeError:
                ok = False
        if not ok:
            bad.append(ch)
    return bad


# ════════════════════════════════════════════════════════════════════════════
#  Turning spreadsheet rows into label "items"
# ════════════════════════════════════════════════════════════════════════════
def load_items(path, profile, limit=None):
    load_fonts()
    notes, items = [], []
    fname = os.path.basename(path)
    headers, rows = read_table(path)
    if not headers:
        return [], [("err", f"{fname}: the spreadsheet is empty.")]
    idx, det = resolve(profile, headers)
    fields, auc = profile["fields"], profile["auction"]

    lost = [FIELD_NAME[k] for k in FIELD_KEYS if fields.get(k) and idx[k] is None]
    lost += ["extra detail “%s”" % short(c, 30) for _, i, c in det if i is None]
    if idx["artist"] is None and idx["title"] is None:
        return [], [("err", f"{fname}: can't find the artist or title columns - this spreadsheet doesn't match the "
                            f"profile “{profile['profile_name']}”.  Run configure to fix or make a new profile.")]
    for name in lost:
        important = name == FIELD_NAME["auction"]
        notes.append(("warn", f"{fname}: couldn't find the column for {name} - "
                              + ("EVERY label would show as NOT for auction, so the auction marker was left off." if important
                                 else "that information is left off the labels.")))

    for rownum, row in rows:
        def g(i):
            return row[i] if i is not None and i < len(row) else None

        artist = clean_text(g(idx["artist"]), multi=False)
        title = clean_text(g(idx["title"]), multi=False)
        if not artist and not title:
            continue
        where = f"{fname}, row {rownum}"
        who = artist or title
        details = []
        for label, i, _ in det:
            v = clean_text(g(i))
            if v and norm(v) not in NA_VALUES:
                details.append((label.strip() + " " + v) if label.strip() else v)
        status = None
        if idx["auction"] is not None:
            ans = clean_text(g(idx["auction"]))
            status = is_yes(ans, auc["yes_words"])
            if not ans:
                notes.append(("warn", f"{where} ({who}): auction question left blank - shown as NOT for auction."))
        if idx["permission"] is not None:
            ans = clean_text(g(idx["permission"]))
            if not is_yes(ans, ["yes"]):
                notes.append(("warn", f"{where} ({who}): display permission isn't a clear Yes ({short(ans or 'blank', 40)}) - please check."))
        item = dict(artist=artist, role=clean_text(g(idx["role"])), title=title, details=details,
                    statement=clean_text(g(idx["statement"]), multi=False), auction=status, where=where)
        bad = unsupported(" ".join([artist, item["role"], title, item["statement"]] + details))
        if bad:
            notes.append(("warn", f"{where} ({who}): contains characters this font can't draw: {' '.join(bad[:8])}"))
        items.append(item)
        if limit and len(items) >= limit:
            break
    return items, notes


# ════════════════════════════════════════════════════════════════════════════
#  Drawing the labels
# ════════════════════════════════════════════════════════════════════════════
def make_styles(size):
    def st(name, font, sz, lead, col=(0, 0, 0)):
        return ParagraphStyle(name, fontName=font, fontSize=sz, leading=lead, textColor=colors.Color(*col))
    return dict(name=st("n", FONTS["bold"], 20, 24), role=st("r", FONTS["reg"], 11, 14, (.4, .4, .4)),
                title=st("t", FONTS["ital"], 16, 20), detail=st("d", FONTS["reg"], 12, 15, (.2, .2, .2)),
                stmt=st("s", FONTS["reg"], size, round(size * 1.28, 2), (.15, .15, .15)))


def _quoted(t):
    return t if re.match(r"^[\"“‘']", t) else "“%s”" % t


class Card:
    """One label's content, measured for a given width and statement font size."""

    def __init__(self, item, w, size, ind, stmt=None, show_stmt=True, compact=False):
        self.item, self.w, self.ind, self.blocks = item, w, ind, []
        iw, S = w - 2 * PAD, make_styles(size)

        def add(kind, para, space, width=None):
            _, h = para.wrap(width or iw, 1e7)
            self.blocks.append((kind, space, para, h))

        if item["artist"]:
            add("name", Paragraph(xml_escape(item["artist"]), S["name"]), 0, iw - (ind["w"] if ind else 0))
        if compact:
            add("role", Paragraph("(continued from the previous page)", S["role"]), 2)
        elif item["role"]:
            add("role", Paragraph(xml_escape(item["role"]), S["role"]), 2)
        if item["title"]:
            add("title", Paragraph(xml_escape(_quoted(item["title"])), S["title"]), 8 if self.blocks else 0)
        if not compact:
            for n, d in enumerate(item["details"]):
                add("detail", Paragraph(xml_escape(d), S["detail"]), 6 if n == 0 else 2)
        if show_stmt:
            if stmt is None and item["statement"]:
                stmt = Paragraph(stmt_markup(item["statement"]), S["stmt"])
            if stmt is not None:
                add("stmt", stmt, RULE_GAP)
        self.h = 2 * PAD + sum(sp + h for _, sp, _, h in self.blocks)


def fit_card(item, w, spref, smin, ind, target_h):
    """Shrink the statement text (never below the minimum) until it fits; otherwise the label just grows."""
    size = spref
    while True:
        card = Card(item, w, size, ind)
        if card.h <= target_h or size <= smin:
            return card
        size = max(smin, size - 0.5)


def draw_card(c, card, x, top, h, foot=None):
    w = card.w
    c.setStrokeColorRGB(.55, .55, .55)
    c.setLineWidth(.75)
    c.rect(x, top - h, w, h, stroke=1, fill=0)
    y, name_top = top - PAD, top - PAD
    for kind, sp, p, ph in card.blocks:
        y -= sp
        if kind == "name":
            name_top = y
        if kind == "stmt":
            c.setStrokeColorRGB(.8, .8, .8)
            c.setLineWidth(.5)
            c.line(x + PAD, y + sp / 2, x + w - PAD, y + sp / 2)
        p.drawOn(c, x + PAD, y - ph)
        y -= ph
    if card.ind:                                  # auction dot (green = for auction, red = not)
        r, cx, cy = 7, x + w - PAD - 7, name_top - 12
        col = (.1, .6, .1) if card.ind["status"] else (.8, .1, .1)
        c.setFillColorRGB(*col)
        c.circle(cx, cy, r, stroke=0, fill=1)
        if card.ind["text"]:
            c.setFont(FONTS["bold"], 10)
            c.drawRightString(cx - r - 8, cy - 3.5, card.ind["text"])
    if foot:
        c.setFillColorRGB(.4, .4, .4)
        c.setFont(FONTS["ital"], 9)
        c.drawRightString(x + w - PAD, top - h + 6, foot)


def render_pdf(items, profile, out_path):
    """Lay the labels out in rows.  A label whose statement is too long simply grows taller; the row grows
    with it and everything after moves down.  A label too tall for a column gets a full-width row, and one
    too tall for a whole page continues onto the next page."""
    load_fonts()
    lay = {**DEFAULT_LAYOUT, **profile["layout"]}
    pw, ph = PAGE_SIZES.get(lay["page_size"], A4)
    pw, ph = (max(pw, ph), min(pw, ph)) if lay["orientation"] == "landscape" else (min(pw, ph), max(pw, ph))
    m = float(lay["margin_mm"]) * mm
    uw, uh = pw - 2 * m, ph - 2 * m
    cols, rows = max(1, int(lay["columns"])), max(1, int(lay["rows"]))
    cw, std_h = uw / cols, uh / rows
    smin = max(MIN_FONT, float(lay["statement_min"]))
    spref = max(smin, float(lay["statement_size"]))
    auc = profile["auction"]

    def indicator(it):
        if it["auction"] is None:
            return None
        text = auc["text_yes"] if it["auction"] else auc["text_no"]
        tw = pdfmetrics.stringWidth(text, FONTS["bold"], 10) if text else 0
        return dict(status=bool(it["auction"]), text=text, w=22 + (tw + 8 if text else 0))

    c = canvas.Canvas(out_path, pagesize=(pw, ph))
    c.setTitle(profile["profile_name"] + " - Labels")
    c.setAuthor("Exhibition Label Studio")
    st = dict(y=0.0, started=False, pages=0, grown=0, wide=0, split=0, labels=0)
    row = []

    def new_page():
        if st["started"]:
            c.showPage()
        st.update(started=True, y=0.0, pages=st["pages"] + 1)

    def ensure_room(h):
        if not st["started"] or st["y"] + h > uh + 0.01:
            new_page()

    def flush():
        if not row:
            return
        h = max(std_h, max(cd.h for cd in row))
        ensure_room(h)
        top = ph - m - st["y"]
        for i, cd in enumerate(row):
            draw_card(c, cd, m + i * cw, top, h)
            st["grown"] += cd.h > std_h + 0.5
        st["y"] += h
        row.clear()

    def split_across_pages(it):
        st["split"] += 1
        iw, ind = uw - 2 * PAD, indicator(it)
        pending = Paragraph(stmt_markup(it["statement"]), make_styles(smin)["stmt"])
        first = True
        while pending is not None:
            head = Card(it, uw, smin, ind if first else None, show_stmt=False, compact=not first)
            room = uh - head.h - RULE_GAP
            pending.wrap(iw, 1e7)
            if pending.height <= room:
                part, rest = pending, None
            else:
                pieces = pending.split(iw, room - 18)
                part, rest = (pieces[0], pieces[1] if len(pieces) > 1 else None)
            card = Card(it, uw, smin, ind if first else None, stmt=part, compact=not first)
            if first and st["started"] and st["y"] == 0:
                pass
            else:
                new_page()
            draw_card(c, card, m, ph - m, uh, foot="continued on the next page  ►" if rest else None)
            st["y"] = uh
            pending, first = rest, False

    for it in items:
        ind = indicator(it)
        card = fit_card(it, cw, spref, smin, ind, std_h)
        if card.h > uh:                                       # too tall for one column
            flush()
            card = fit_card(it, uw, spref, smin, ind, std_h) if cols > 1 else card
            if card.h > uh:
                split_across_pages(it)
                st["labels"] += 1
                continue
            h = max(std_h, card.h)
            ensure_room(h)
            draw_card(c, card, m, ph - m - st["y"], h)
            st["y"] += h
            st["wide"] += cols > 1
            st["grown"] += card.h > std_h + 0.5
            st["labels"] += 1
            continue
        row.append(card)
        st["labels"] += 1
        if len(row) == cols:
            flush()
    flush()
    if st["started"]:
        c.save()
    return st


def build_labels(profile, files, out_path, limit=None):
    """Read every file, tidy, de-duplicate, sort and draw.  Returns (stats or None, notes)."""
    items, notes = [], []
    for f in files:
        try:
            its, nts = load_items(f, profile, limit)
        except Exception as e:
            notes.append(("err", f"Couldn't read {os.path.basename(f)}: {e}"))
            continue
        items += its
        notes += nts
    seen, unique = set(), []
    for it in items:
        key = (norm(it["artist"]), norm(it["title"]), norm(it["statement"]))
        if key in seen:
            notes.append(("info", f"{it['where']}: identical entry already included - skipped the duplicate."))
            continue
        seen.add(key)
        unique.append(it)
    sort = profile["layout"].get("sort", "file")
    if sort == "artist":
        unique.sort(key=lambda i: norm(i["artist"]))
    elif sort == "title":
        unique.sort(key=lambda i: norm(i["title"]))
    if limit:
        unique = unique[:limit]
    if not unique:
        return None, notes
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    try:
        stats = render_pdf(unique, profile, out_path)
    except PermissionError:                                   # PDF is open in a viewer
        root, ext = os.path.splitext(out_path)
        out_path = f"{root} ({datetime.now():%H-%M-%S}){ext}"
        notes.append(("warn", "The old PDF was open in another program, so a new copy was saved beside it."))
        stats = render_pdf(unique, profile, out_path)
    stats["path"] = out_path
    return stats, notes


# ════════════════════════════════════════════════════════════════════════════
#  Terminal look & feel  (plain Python - no extra packages, Windows + Linux)
# ════════════════════════════════════════════════════════════════════════════
IS_TTY = sys.stdin.isatty() and sys.stdout.isatty()
COLOR = sys.stdout.isatty() and "NO_COLOR" not in os.environ
ACC, VIO, DIM = "38;5;80", "38;5;141", "38;5;245"
GOOD, WARN, BAD, BOLD, REV = "38;5;78", "38;5;221", "38;5;203", "1", "7"
SELECT = "48;5;61;38;5;231;1"
_ANSI = re.compile(r"\x1b\[[0-9;?]*[A-Za-z]")


def sty(text, *codes):
    return "\x1b[%sm%s\x1b[0m" % (";".join(codes), text) if COLOR and codes else str(text)


def vlen(s):
    return len(_ANSI.sub("", s))


def term_size():
    s = shutil.get_terminal_size((100, 30))
    return s.columns, s.lines


def W():
    return max(50, min(term_size()[0] - 2, 92))


def wrap_text(text, width):
    out = []
    for para in str(text).split("\n"):
        out += textwrap.wrap(para, max(10, width)) or [""]
    return out


def enable_vt():
    if IS_WIN:
        try:
            import ctypes
            k = ctypes.windll.kernel32
            h, mode = k.GetStdHandle(-11), ctypes.c_uint32()
            if k.GetConsoleMode(h, ctypes.byref(mode)):
                k.SetConsoleMode(h, mode.value | 0x0004)
        except Exception:
            pass


def banner():
    tw, th = term_size()
    if th < 24 or tw < 70:
        return ["  " + sty("░▒▓█ ", VIO) + sty("EXHIBITION LABEL STUDIO", BOLD, ACC) + sty("  gallery labels made easy", DIM)]
    w = min(tw - 2, 76)

    def row(content):
        return "  " + sty("║", VIO) + content + " " * (w - 2 - vlen(content)) + sty("║", VIO)
    title = "  " + sty("░▒▓█", VIO) + "  " + sty("E X H I B I T I O N   L A B E L   S T U D I O", BOLD, ACC) + "  " + sty("█▓▒░", VIO)
    return ["  " + sty("╔" + "═" * (w - 2) + "╗", VIO), row(title),
            row("       " + sty("Beautiful gallery labels, made from your entry form", DIM)),
            "  " + sty("╚" + "═" * (w - 2) + "╝", VIO)]


# ── keyboard ────────────────────────────────────────────────────────────────
_TUI = {"on": False, "old": None}
_SEQ = {b"[A": "up", b"OA": "up", b"[B": "down", b"OB": "down", b"[C": "right", b"OC": "right", b"[D": "left",
        b"OD": "left", b"[H": "home", b"OH": "home", b"[1~": "home", b"[F": "end", b"OF": "end", b"[4~": "end",
        b"[3~": "delete", b"[5~": "pgup", b"[6~": "pgdn"}


def tui_start():
    if _TUI["on"]:
        return
    enable_vt()
    if not IS_WIN:
        _TUI["old"] = termios.tcgetattr(0)
        tty.setcbreak(0)
    sys.stdout.write("\x1b[?1049h\x1b[?25l")
    sys.stdout.flush()
    _TUI["on"] = True


def tui_stop():
    if not _TUI["on"]:
        return
    sys.stdout.write("\x1b[?25h\x1b[?1049l")
    sys.stdout.flush()
    if not IS_WIN and _TUI["old"]:
        termios.tcsetattr(0, termios.TCSADRAIN, _TUI["old"])
    _TUI["on"] = False


atexit.register(tui_stop)


def read_key():
    if IS_WIN:
        ch = msvcrt.getwch()
        if ch in ("\x00", "\xe0"):
            return {"H": "up", "P": "down", "K": "left", "M": "right", "G": "home", "O": "end", "I": "pgup",
                    "Q": "pgdn", "S": "delete"}.get(msvcrt.getwch(), "")
        return {"\r": "enter", "\x1b": "esc", "\x08": "backspace", "\x7f": "backspace", "\t": "tab"}.get(ch, ch)
    fd = sys.stdin.fileno()
    b = os.read(fd, 1)
    if not b:
        raise EOFError
    if b == b"\x1b":
        if not select.select([fd], [], [], 0.05)[0]:
            return "esc"
        seq = os.read(fd, 1)
        if seq in (b"[", b"O"):
            while select.select([fd], [], [], 0.05)[0]:
                ch = os.read(fd, 1)
                seq += ch
                if 0x40 <= ch[0] <= 0x7E:
                    break
        return _SEQ.get(seq, "")
    if b in (b"\r", b"\n"):
        return "enter"
    if b in (b"\x7f", b"\x08"):
        return "backspace"
    if b == b"\t":
        return "tab"
    if b[0] >= 0xC0:
        need = 1 if b[0] < 0xE0 else 2 if b[0] < 0xF0 else 3
        while need:
            b += os.read(fd, 1)
            need -= 1
    return b.decode("utf-8", "replace")


def paint(lines):
    _, th = term_size()
    sys.stdout.write("\x1b[H" + "".join(l + "\x1b[K\n" for l in lines[: th - 1]) + "\x1b[J")
    sys.stdout.flush()


# ── widgets ─────────────────────────────────────────────────────────────────
def menu(title, items, subtitle="", crumb="", default=0, esc="Back"):
    """Arrow-key menu.  items: dicts with label, desc, tag, tag_style, sep.  Returns chosen index or None (Esc)."""
    pick = [k for k, it in enumerate(items) if not it.get("sep")]
    if not pick:
        return None
    i = default if 0 <= default < len(items) and not items[default].get("sep") else pick[0]
    top = 0
    while True:
        tw, th = term_size()
        w = W()
        head = banner() + [""]
        if crumb:
            head.append("  " + sty(crumb, DIM))
        head += ["  " + sty(l, BOLD, ACC) for l in wrap_text(title, w - 4)]
        head += ["  " + l for l in wrap_text(subtitle, w - 4)] if subtitle else []
        head.append("")
        desc = wrap_text(items[i].get("desc", ""), w - 8)[:4] if items[i].get("desc") else []
        foot = [""] + ["   " + sty("│ ", VIO) + sty(l, DIM) for l in desc] + [""]
        foot.append("  " + sty("↑ ↓  choose      Enter  select      Esc  " + esc.lower(), DIM))
        vis = max(3, min(len(items), th - len(head) - len(foot) - 2))
        top = min(max(top, i - vis + 1), i) if i >= 0 else 0
        top = max(0, min(top, len(items) - vis))
        tagw = min(34, max([len(it.get("tag", "")) for it in items] + [0]))
        lw = w - 8 - (tagw + 2 if tagw else 0)
        body = []
        for k in range(top, min(len(items), top + vis)):
            it = items[k]
            if it.get("sep"):
                body.append("  " + sty(("── " + it.get("label", "") + " ").ljust(w - 4, "─"), DIM))
                continue
            label = short(it["label"], lw).ljust(lw)
            tag = short(it.get("tag", ""), tagw).rjust(tagw)
            if k == i:
                body.append(" " + sty(" ►  %s%s " % (label, ("  " + tag) if tagw else ""), SELECT))
            else:
                body.append("     " + label + (("  " + sty(tag, *(it.get("tag_style") or (DIM,)))) if tagw else ""))
        if len(items) > vis:
            foot[-1] += sty("      (%d of %d)" % (pick.index(i) + 1 if i in pick else 0, len(pick)), DIM)
        paint(head + body + foot)
        key = read_key()
        step = {"up": -1, "down": 1}.get(key)
        if step:
            j = i
            while True:
                j = (j + step) % len(items)
                if not items[j].get("sep"):
                    break
            i = j
        elif key == "home":
            i = pick[0]
        elif key == "end":
            i = pick[-1]
        elif key in ("pgup", "pgdn"):
            j = max(0, min(len(items) - 1, i + (vis if key == "pgdn" else -vis)))
            i = min(pick, key=lambda p: abs(p - j))
        elif key == "enter":
            return i
        elif key == "esc":
            return None


def ask_text(title, prompt, default="", help="", crumb="", validate=None):
    """Single-line text box.  Returns the text, or None if Esc."""
    buf, pos, err = list(default), len(default), ""
    while True:
        w = W()
        avail = w - 12
        start = max(0, pos - avail + 1)
        vis = buf[start:start + avail]
        p = pos - start
        text = "".join(vis[:p]) + sty(vis[p] if p < len(vis) else " ", REV) + "".join(vis[p + 1:])
        lines = banner() + [""]
        if crumb:
            lines.append("  " + sty(crumb, DIM))
        lines += ["  " + sty(l, BOLD, ACC) for l in wrap_text(title, w - 4)]
        lines += ["  " + l for l in wrap_text(help, w - 4)] if help else []
        lines += ["", "  " + prompt, "  " + sty("┌" + "─" * (w - 6) + "┐", VIO),
                  "  " + sty("│ ", VIO) + text + " " * (w - 8 - min(len(vis) + 1, avail)) + sty(" │", VIO),
                  "  " + sty("└" + "─" * (w - 6) + "┘", VIO)]
        if err:
            lines += ["", "  " + sty("× " + err, BAD)]
        lines += ["", "  " + sty("Type your answer     Enter  confirm      Esc  cancel", DIM)]
        paint(lines)
        k = read_key()
        if k == "enter":
            val = "".join(buf).strip()
            err = validate(val) if validate else ""
            if not err:
                return val
        elif k == "esc":
            return None
        elif k == "backspace" and pos:
            del buf[pos - 1]
            pos -= 1
        elif k == "delete" and pos < len(buf):
            del buf[pos]
        elif k == "left":
            pos = max(0, pos - 1)
        elif k == "right":
            pos = min(len(buf), pos + 1)
        elif k == "home":
            pos = 0
        elif k == "end":
            pos = len(buf)
        elif len(k) == 1 and k.isprintable():
            buf.insert(pos, k)
            pos += 1


def notice(kind, title, lines, crumb=""):
    col = {"ok": GOOD, "warn": WARN, "err": BAD, "info": ACC}[kind]
    w = W()
    body = []
    for l in lines:
        body += ["   " + x for x in wrap_text(l, w - 8)]
    paint(banner() + [""] + (["  " + sty(crumb, DIM)] if crumb else []) + ["  " + sty(title, BOLD, col), ""] + body +
          ["", "  " + sty("Press Enter to continue", DIM)])
    while read_key() not in ("enter", "esc", " "):
        pass


def confirm(title, question, yes="Yes", no="No", default_yes=True, crumb=""):
    r = menu(title, [dict(label=yes), dict(label=no)], subtitle=question, crumb=crumb, default=0 if default_yes else 1)
    return None if r is None else r == 0


def choose(title, options, current, subtitle="", crumb=""):
    """options = [(label, value, description)] -> value or None"""
    items = [dict(label=l, desc=d, tag="● current" if v == current else "", tag_style=(GOOD,)) for l, v, d in options]
    default = next((k for k, (_, v, _) in enumerate(options) if v == current), 0)
    r = menu(title, items, subtitle=subtitle, crumb=crumb, default=default, esc="Cancel")
    return None if r is None else options[r][1]


def open_path(path):
    try:
        if IS_WIN:
            os.startfile(path)
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


# ════════════════════════════════════════════════════════════════════════════
#  Output for "make the labels" (ordinary scrolling text, still pretty)
# ════════════════════════════════════════════════════════════════════════════
def say(kind, text, indent=2):
    mark = {"ok": sty("●", GOOD), "info": sty("●", ACC), "warn": sty("▲", WARN), "err": sty("×", BAD),
            "step": sty("►", VIO)}[kind]
    print(" " * indent + mark + " " + text)


def box(lines, color):
    w = max(vlen(l) for l in lines) + 4
    print("  " + sty("┌" + "─" * w + "┐", color))
    for l in lines:
        print("  " + sty("│", color) + "  " + l + " " * (w - 2 - vlen(l)) + sty("│", color))
    print("  " + sty("└" + "─" * w + "┘", color))


def run_generation(profile, files, limit=None, out_path=None):
    say("step", "Using profile: " + sty(profile["profile_name"], BOLD))
    lay = profile["layout"]
    say("info", "Layout: %s × %s labels per page · %s %s · statement text %g–%g pt"
        % (lay["columns"], lay["rows"], lay["page_size"], lay["orientation"],
           max(MIN_FONT, float(lay["statement_min"])), max(float(lay["statement_min"]), float(lay["statement_size"]))))
    out = out_path or os.path.join(OUTPUT_FOLDER, safe_filename(profile["profile_name"]) + " - Labels.pdf")
    say("step", "Building labels from %d file%s..." % (len(files), "" if len(files) == 1 else "s"))
    stats, notes = build_labels(profile, files, out, limit)
    print()
    errs = [t for k, t in notes if k == "err"]
    warns = [t for k, t in notes if k == "warn"]
    infos = [t for k, t in notes if k == "info"]
    for t in errs:
        say("err", t)
    for t in warns[:15]:
        say("warn", t)
    if len(warns) > 15:
        say("warn", "...and %d more warnings." % (len(warns) - 15))
    for t in infos[:5]:
        say("info", t)
    if errs or warns or infos:
        print()
    if not stats:
        box([sty("No labels were made.", BAD),
             "Check the messages above, or run configure to fix the profile."], BAD)
        return None
    extra = []
    if stats["grown"]:
        extra.append("%d label%s grew taller to fit a long statement" % (stats["grown"], "" if stats["grown"] == 1 else "s"))
    if stats["wide"]:
        extra.append("%d very long one%s got a full-width label" % (stats["wide"], "" if stats["wide"] == 1 else "s"))
    if stats["split"]:
        extra.append("%d extremely long one%s continue%s on a second page" % (stats["split"], "" if stats["split"] == 1 else "s", "s" if stats["split"] == 1 else ""))
    lines = [sty("Done!  %d labels on %d page%s" % (stats["labels"], stats["pages"], "" if stats["pages"] == 1 else "s"), BOLD, GOOD)]
    lines += [sty(e, DIM) for e in extra]
    lines += ["", "Saved to:", sty(os.path.relpath(stats["path"], BASE), ACC)]
    box(lines, GOOD)
    return stats


def pick_profile_for_generate(profiles, heads, wanted):
    if wanted:
        for p in profiles:
            if norm(p["profile_name"]) == norm(wanted):
                return p
        say("err", "No profile called “%s”.  Available: %s" % (wanted, ", ".join(p["profile_name"] for p in profiles)))
        return None
    scores = {p["profile_name"]: (sum(match_score(p, h) for h in heads.values()) / len(heads)) if heads else 0.0
              for p in profiles}
    best = max(range(len(profiles)), key=lambda k: scores[profiles[k]["profile_name"]])
    if len(profiles) == 1 or not IS_TTY:
        return profiles[best]
    tui_start()
    try:
        items = [dict(label=p["profile_name"], tag="%d%% match" % round(scores[p["profile_name"]] * 100),
                      tag_style=(GOOD if scores[p["profile_name"]] > .8 else WARN,),
                      desc="How well this profile's columns fit your spreadsheet(s).") for p in profiles]
        r = menu("Which form profile should be used?", items, default=best, esc="Quit",
                 subtitle="You have more than one profile.  The best match is highlighted.")
    finally:
        tui_stop()
    return None if r is None else profiles[r]


def cmd_generate(args):
    enable_vt()
    print()
    print("\n".join(banner()))
    print()
    os.makedirs(INPUT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    profiles = load_profiles()
    for bad in BAD_PROFILES:
        say("warn", "Skipped unreadable profile file: " + bad)
    if not profiles:
        say("err", "There are no form profiles yet.  Run the configure script to create one.")
        return 1
    files = find_data_files()
    if not files:
        say("warn", "No spreadsheets found.  Drop your .xlsx or .csv files into the folder:")
        say("info", sty("EXCEL_FILES_HERE", ACC, BOLD) + "   and run this again.")
        return 1
    say("ok", "Found %d spreadsheet%s in EXCEL_FILES_HERE" % (len(files), "" if len(files) == 1 else "s"))
    heads = {}
    for f in files:
        try:
            heads[f] = read_table(f)[0]
        except Exception:
            pass
    prof = pick_profile_for_generate(profiles, heads, args.profile)
    if prof is None:
        return 1
    return 0 if run_generation(prof, files, args.limit) else 1


# ════════════════════════════════════════════════════════════════════════════
#  Configure  (the friendly set-up screens)
# ════════════════════════════════════════════════════════════════════════════
SESSION = {"headers": [], "rows": []}


def read_example(path):
    try:
        headers, rows = read_table(path)
    except Exception as e:
        notice("err", "That file couldn't be opened", [str(e), "Make sure it is a .xlsx or .csv file and is not open in another program."])
        return None
    if not headers:
        notice("err", "That spreadsheet looks empty", ["Export the form responses to Excel again and try once more."])
        return None
    SESSION.update(headers=headers, rows=rows)
    return headers, rows


def pick_data_file(title, crumb=""):
    files = find_data_files()
    items = [dict(label=os.path.basename(f), desc="Found in the EXCEL_FILES_HERE folder.", value=f) for f in files]
    items.append(dict(label="Type or paste a file location…", value="__path__",
                      desc="Tip: on most computers you can drag a file into this window to paste its location."))
    r = menu(title, items, crumb=crumb, esc="Cancel",
             subtitle="Choose a spreadsheet exported from your form.  It is only used to read the column names."
             if files else "No spreadsheets are in the EXCEL_FILES_HERE folder yet, so you can paste a location instead.")
    if r is None:
        return None
    if items[r]["value"] != "__path__":
        return items[r]["value"]
    while True:
        s = ask_text("Where is the spreadsheet?", "File location:", crumb=crumb,
                     help="Example:  C:\\Users\\me\\Downloads\\responses.xlsx   or   /home/me/responses.xlsx")
        if s is None:
            return None
        s = os.path.expanduser(s.strip().strip("\"'").replace("\\ ", " "))
        if os.path.isfile(s) and s.lower().endswith(DATA_EXTS):
            return s
        notice("err", "I couldn't use that", ["The file wasn't found, or it isn't an .xlsx / .csv file."])


def sample_value(headers, i):
    if SESSION["headers"] != headers:
        return ""
    for _, row in SESSION["rows"][:10]:
        v = clean_text(row[i]) if i < len(row) else ""
        if v:
            return v
    return ""


def pick_column(prof, key, headers, crumb, optional, suggestion=None, question=""):
    """Column chooser.  Returns column index, -1 for 'leave out', or None for Esc."""
    taken = {k: find_col(headers, v) for k, v in prof["fields"].items() if v and k != key}
    used = {i: FIELD_NAME[k] for k, i in taken.items() if i is not None}
    cur = find_col(headers, prof["fields"].get(key, "")) if key in prof["fields"] else None
    items = []
    if optional:
        items.append(dict(label="— Leave this out —", value=-1, tag="skip", desc="This form doesn't ask for it, or you don't want it on the labels."))
    for i, h in enumerate(headers):
        tag, style = "", (DIM,)
        if i == cur:
            tag, style = "● current", (GOOD,)
        elif i == suggestion:
            tag, style = "★ suggested", (WARN,)
        elif i in used:
            tag = "in use: " + used[i]
        ex = sample_value(headers, i)
        full = re.sub(r"\s+", " ", h.replace("\xa0", " ")).strip()
        desc = ("Example answer:  “%s”" % short(ex, 150) if ex else "") + ("\n" if ex else "") + short(full, 200)
        items.append(dict(label=nice_header(h, 70) or "(blank)", tag=tag, tag_style=style, value=i, desc=desc))
    default = next((k for k, it in enumerate(items) if it["value"] == (cur if cur is not None else suggestion)), 0)
    r = menu(question, items, crumb=crumb, default=default, esc="Cancel",
             subtitle="Pick the column from your spreadsheet.  ★ marks my best guess.")
    return None if r is None else items[r]["value"]


def set_field(prof, key, headers, crumb, optional, suggest=True):
    """Ask for one field.  Returns False if the user pressed Esc."""
    q = next(x[2] for x in FIELDS if x[0] == key)
    guess = guess_column(headers, next(x[4] for x in FIELDS if x[0] == key)) if suggest else None
    r = pick_column(prof, key, headers, crumb, optional, guess, q)
    if r is None:
        return False
    prof["fields"][key] = "" if r == -1 else header_phrase(headers[r])
    return True


def edit_details(prof, headers, crumb):
    while True:
        items = []
        for n, d in enumerate(prof["details"]):
            j = find_col(headers, d["column"])
            items.append(dict(label="Line %d:  %s‹%s›" % (n + 1, (d["label"] + " ") if d.get("label") else "", nice_header(headers[j], 40) if j is not None else d["column"]),
                              tag="" if j is not None else "not found", tag_style=(WARN,), value=("edit", n),
                              desc="Press Enter to change or remove this line."))
        items.append(dict(label="+ Add a line", value=("add", 0), desc="For example the MEDIUM, the SIZE, the PRICE… anything you'd like printed under the title."))
        items.append(dict(label="Done", value=("done", 0)))
        r = menu("Extra detail lines", items, crumb=crumb, default=len(items) - 1, esc="Done",
                 subtitle="These lines are printed under the artwork title, in order.")
        if r is None or items[r]["value"][0] == "done":
            return
        kind, n = items[r]["value"]
        if kind == "add":
            tmp = new_profile("tmp")
            c = pick_column(tmp, "_", headers, crumb, False, None, "Which column should this line show?")
            if c is None:
                continue
            lab = ask_text("Text before the answer", "Label (optional):", crumb=crumb,
                           help="For example  Medium:  or  Size:  - or leave it empty to print just the answer.")
            if lab is None:
                continue
            prof["details"].append({"label": lab, "column": header_phrase(headers[c])})
        else:
            d = prof["details"][n]
            a = choose("What would you like to do?", [("Change the column", "col", ""), ("Change the label text", "lab", ""),
                                                        ("Move up", "up", ""), ("Remove this line", "del", "")], None, crumb=crumb)
            if a == "col":
                tmp = new_profile("tmp")
                c = pick_column(tmp, "_", headers, crumb, False, None, "Which column should this line show?")
                if c is not None:
                    d["column"] = header_phrase(headers[c])
            elif a == "lab":
                lab = ask_text("Text before the answer", "Label (optional):", d.get("label", ""), crumb=crumb)
                if lab is not None:
                    d["label"] = lab
            elif a == "up" and n > 0:
                prof["details"][n - 1], prof["details"][n] = prof["details"][n], prof["details"][n - 1]
            elif a == "del":
                del prof["details"][n]


def layout_menu(prof, crumb):
    L = prof["layout"]
    while True:
        smin = max(MIN_FONT, float(L["statement_min"]))
        items = [
            dict(label="Paper size", tag=L["page_size"], value="page_size"),
            dict(label="Paper direction", tag=L["orientation"].capitalize(), value="orientation"),
            dict(label="Labels across the page", tag=str(L["columns"]), value="columns",
                 desc="More columns = narrower labels."),
            dict(label="Labels down the page", tag=str(L["rows"]), value="rows",
                 desc="This is the standard label height.  Labels with long statements grow taller than this."),
            dict(label="Statement text size", tag="%g pt" % L["statement_size"], value="statement_size",
                 desc="Preferred size.  If a statement doesn't fit, the text shrinks - but never below the smallest size below."),
            dict(label="Smallest statement text allowed", tag="%g pt" % smin, value="statement_min",
                 desc="Text never goes below 10 pt.  Past this, the label grows taller instead."),
            dict(label="Margin around the page", tag="%g mm" % L["margin_mm"], value="margin_mm",
                 desc="Printers can't print right to the edge of the paper, so keep at least 5 mm."),
            dict(label="Order of the labels", tag={"file": "as in spreadsheet", "artist": "artist A–Z", "title": "title A–Z"}[L["sort"]], value="sort"),
            dict(label="Auction wording", tag="“%s”" % prof["auction"]["text_yes"], value="auction",
                 desc="The words printed beside the green dot.  (The red dot can carry words too.)"),
            dict(label="Back", value="back"),
        ]
        r = menu("Label layout & style", items, crumb=crumb, esc="Back")
        if r is None or items[r]["value"] == "back":
            return
        k = items[r]["value"]
        opts = {
            "page_size": [(n, n, "") for n in PAGE_SIZES],
            "orientation": [("Landscape (wide)", "landscape", ""), ("Portrait (tall)", "portrait", "")],
            "columns": [(str(n), n, "") for n in (1, 2, 3)],
            "rows": [(str(n), n, "") for n in (1, 2, 3, 4)],
            "statement_size": [("%d pt" % n, n, "") for n in range(10, 19)],
            "statement_min": [("%d pt" % n, n, "") for n in (10, 11, 12, 14)],
            "margin_mm": [("%d mm" % n, n, "") for n in (5, 8, 10, 12, 15, 20)],
            "sort": [("As in the spreadsheet", "file", ""), ("Artist name A–Z", "artist", ""), ("Artwork title A–Z", "title", "")],
        }
        if k == "auction":
            t = ask_text("Auction wording", "Words beside the GREEN dot:", prof["auction"]["text_yes"], crumb=crumb,
                         help="Leave empty to show just the dot.")
            if t is not None:
                prof["auction"]["text_yes"] = t
            t = ask_text("Auction wording", "Words beside the RED dot:", prof["auction"]["text_no"], crumb=crumb,
                         help="Leave empty to show just a red dot (or type something like:  Not for auction).")
            if t is not None:
                prof["auction"]["text_no"] = t
        else:
            v = choose(items[r]["label"], opts[k], L[k], crumb=crumb)
            if v is not None:
                L[k] = v


def make_test_pdf(prof, crumb):
    path = pick_data_file("Make a test PDF - pick a spreadsheet", crumb)
    if not path:
        return
    out = os.path.join(OUTPUT_FOLDER, "PREVIEW - " + safe_filename(prof["profile_name"]) + ".pdf")
    stats, notes = build_labels(prof, [path], out, limit=6)
    if not stats:
        notice("err", "Nothing could be made", [t for _, t in notes] or ["No usable rows were found."], crumb)
        return
    lines = ["A test PDF with the first %d labels was saved to:" % stats["labels"], os.path.relpath(stats["path"], BASE), ""]
    lines += ["▲ " + t for k, t in notes if k in ("warn", "err")][:6]
    lines += ["", "It should open automatically.  If it doesn't, open the Generated_Labels folder."]
    open_path(stats["path"])
    notice("ok", "Test PDF ready", lines, crumb)


def ensure_headers(prof, crumb):
    """The saved form columns - or ask for an example spreadsheet if none are saved yet."""
    if prof["form_columns"]:
        return prof["form_columns"]
    notice("info", "One quick thing first", ["This profile doesn't have a list of the form's columns saved yet.",
                                             "Pick an example spreadsheet from the form so I can show you the columns."], crumb)
    path = pick_data_file("Pick an example spreadsheet", crumb)
    got = read_example(path) if path else None
    if not got:
        return None
    prof["form_columns"] = got[0]
    save_profile(prof)
    return got[0]


def profile_hub(prof):
    while True:
        heads = prof["form_columns"]
        crumb = "Profiles  ›  " + prof["profile_name"]
        items = []
        for k, name, q, req, _ in FIELDS:
            phrase = prof["fields"].get(k, "")
            if not phrase:
                tag, style = "(not used)", (DIM,)
            elif heads and find_col(heads, phrase) is None:
                tag, style = "▲ not in form", (WARN,)
            else:
                j = find_col(heads, phrase) if heads else None
                tag, style = (nice_header(heads[j], 34) if j is not None else short(phrase, 34)), (GOOD,)
            items.append(dict(label=name + (" *" if req else ""), tag=tag, tag_style=style, value=("field", k),
                              desc=q + "   Press Enter to change it."))
        items.append(dict(label="Extra detail lines", tag="%d line%s" % (len(prof["details"]), "" if len(prof["details"]) == 1 else "s"),
                          value=("details", 0), desc="Medium, size, price… printed under the title."))
        L = prof["layout"]
        items += [
            dict(sep=True, label="Look"),
            dict(label="Label layout & style", tag="%s %s · %s×%s" % (L["page_size"], L["orientation"], L["columns"], L["rows"]), value=("layout", 0),
                 desc="Paper size, labels per page, text size, ordering, auction wording."),
            dict(label="Make a test PDF", value=("test", 0), desc="Builds the first few labels from a spreadsheet so you can check how they look."),
            dict(sep=True, label="Manage"),
            dict(label="Read columns from a new spreadsheet", value=("rescan", 0), desc="Use this if the form gained new questions."),
            dict(label="Rename this profile", value=("rename", 0)),
            dict(label="Make a copy of this profile", value=("copy", 0), desc="Handy for next year's exhibition."),
            dict(label="Delete this profile", value=("delete", 0)),
            dict(label="Back to the main menu", value=("back", 0)),
        ]
        r = menu("Profile: " + prof["profile_name"], items, crumb="Profiles", esc="Back",
                 subtitle="Everything is saved automatically.  * = needed on every label.")
        if r is None or items[r]["value"][0] == "back":
            return
        kind, k = items[r]["value"]
        if kind == "field":
            h = ensure_headers(prof, crumb)
            if h:
                req = next(x[3] for x in FIELDS if x[0] == k)
                set_field(prof, k, h, crumb, optional=not req, suggest=False)
        elif kind == "details":
            h = ensure_headers(prof, crumb)
            if h:
                edit_details(prof, h, crumb)
        elif kind == "layout":
            layout_menu(prof, crumb)
        elif kind == "test":
            make_test_pdf(prof, crumb)
        elif kind == "rescan":
            path = pick_data_file("Read columns from a spreadsheet", crumb)
            got = read_example(path) if path else None
            if got:
                prof["form_columns"] = got[0]
        elif kind == "rename":
            n = ask_text("Rename profile", "New name:", prof["profile_name"], crumb=crumb, validate=lambda v: name_problem(v, prof))
            if n:
                prof["profile_name"] = n
        elif kind == "copy":
            n = ask_text("Copy profile", "Name for the copy:", prof["profile_name"] + " (copy)", crumb=crumb, validate=lambda v: name_problem(v, None))
            if n:
                dup = copy.deepcopy(prof)
                dup["profile_name"] = n
                dup.pop("_file", None)
                save_profile(dup)
                notice("ok", "Copy created", ["“%s” is ready.  Choose it from the profile list to edit it." % n], crumb)
        elif kind == "delete":
            if confirm("Delete this profile?", "“%s” will be removed for good.  (Your spreadsheets are not touched.)" % prof["profile_name"],
                       yes="Yes, delete it", no="No, keep it", default_yes=False, crumb=crumb):
                if prof.get("_file") and os.path.exists(prof["_file"]):
                    os.remove(prof["_file"])
                return
        save_profile(prof)


def name_problem(v, current):
    if not v:
        return "Please type a name."
    if not safe_filename(v).strip():
        return "That name can't be used."
    p = profile_path(v)
    if os.path.exists(p) and not (current and current.get("_file") and os.path.abspath(current["_file"]) == os.path.abspath(p)):
        return "A profile with that name already exists."
    return ""


def wizard_new():
    crumb = "New profile"
    notice("info", "Let's teach the app about your form", [
        "A profile remembers what each column of your form means - which one is the artist's name, which is the title, and so on.",
        "You only do this once per form.  I'll guess the answers for you, so most steps are just a matter of pressing Enter.",
        "You'll need an example spreadsheet exported from the form (Forms → Open in Excel).",
    ], crumb)
    path = pick_data_file("Step 1 - Pick an example spreadsheet", crumb + "  ›  Step 1 of 4")
    if not path:
        return
    got = read_example(path)
    if not got:
        return
    headers, rows = got
    name = ask_text("Step 2 - Name this form profile", "Profile name:", "", crumb=crumb + "  ›  Step 2 of 4",
                    help="Something like:  What Orana Means To Me - 2026   (exhibition name, a dash, then the year).",
                    validate=lambda v: name_problem(v, None))
    if not name:
        return
    prof = new_profile(name)
    prof["form_columns"] = headers
    order = [k for k, *_ in FIELDS]
    for n, key in enumerate(order, 1):
        req = next(x[3] for x in FIELDS if x[0] == key)
        if not set_field(prof, key, headers, "%s  ›  Step 3 of 4  ›  Question %d of %d" % (crumb, n, len(order)), optional=not req):
            if confirm("Stop setting up?", "Your answers so far will be thrown away.", yes="Yes, stop", no="No, carry on",
                       default_yes=False, crumb=crumb) in (True, None):
                return
            set_field(prof, key, headers, crumb, optional=not req)
    used = {find_col(headers, v) for v in prof["fields"].values() if v}
    for hints in (["medium", "material"], ["dimension", "size"]):
        g = guess_column(headers, hints, taken=used)
        if g is not None:
            used.add(g)
            prof["details"].append({"label": "", "column": header_phrase(headers[g])})
    edit_details(prof, headers, crumb + "  ›  Step 4 of 4")
    save_profile(prof)
    if confirm("Profile saved!", "“%s” is ready to use.  Would you like to make a test PDF now to see how it looks?" % name,
               yes="Yes, make a test PDF", no="Not now", crumb=crumb):
        make_test_pdf(prof, crumb)
    profile_hub(prof)


def folders_menu():
    r = choose("Open a folder", [("Where I drop my spreadsheets  (EXCEL_FILES_HERE)", "in", ""),
                                 ("Where the finished labels appear  (Generated_Labels)", "out", ""),
                                 ("Where the profiles are kept  (profiles)", "pro", "")], None)
    if r:
        path = {"in": INPUT_FOLDER, "out": OUTPUT_FOLDER, "pro": PROFILES_DIR}[r]
        os.makedirs(path, exist_ok=True)
        open_path(path)


def make_labels_now(profiles):
    files = find_data_files()
    if not files:
        notice("warn", "No spreadsheets yet", ["Drop your .xlsx or .csv files into the EXCEL_FILES_HERE folder, then try again.",
                                               "(Use “Open my folders” on the main menu to jump straight there.)"])
        return
    items = [dict(label=p["profile_name"]) for p in profiles]
    r = menu("Make labels with which profile?", items, esc="Cancel")
    if r is None:
        return
    tui_stop()
    print()
    say("ok", "Found %d spreadsheet%s" % (len(files), "" if len(files) == 1 else "s"))
    stats = run_generation(profiles[r], files)
    if stats:
        open_path(stats["path"])
    print()
    try:
        input("  Press Enter to go back to the menu...")
    except EOFError:
        pass
    tui_start()


def cmd_configure(args):
    if not IS_TTY:
        print("Please run the configure script from a normal terminal / command window.")
        return 1
    tui_start()
    try:
        while True:
            profiles = load_profiles()
            items = [dict(label="Create a new profile", value="new",
                          desc="Step-by-step: teach the app about a new entry form.  Best for a new exhibition or a different school."),
                     dict(label="Edit a profile", value="edit", tag="%d saved" % len(profiles),
                          desc="Change which columns are used, the look of the labels, or rename / copy / delete a profile."),
                     dict(label="Make labels now", value="make",
                          desc="Builds the PDF from the spreadsheets in EXCEL_FILES_HERE."),
                     dict(label="Open my folders", value="folders", desc="Jump to the folders where spreadsheets, profiles and finished labels live."),
                     dict(label="Quit", value="quit")]
            for bad in BAD_PROFILES:
                items.insert(2, dict(label="(unreadable file: %s)" % bad, sep=True))
            r = menu("What would you like to do?", items, esc="Quit",
                     subtitle="Use the ↑ ↓ arrow keys and press Enter.  Nothing you do here can break your spreadsheets.")
            if r is None or items[r].get("value") == "quit":
                break
            v = items[r]["value"]
            if v == "new":
                wizard_new()
            elif v == "edit":
                if not profiles:
                    notice("info", "No profiles yet", ["Choose “Create a new profile” first."])
                    continue
                q = menu("Which profile?", [dict(label=p["profile_name"], tag="%d columns" % len([x for x in p["fields"].values() if x])) for p in profiles],
                         crumb="Edit a profile", esc="Back")
                if q is not None:
                    profile_hub(profiles[q])
            elif v == "make":
                if not profiles:
                    notice("info", "No profiles yet", ["Choose “Create a new profile” first."])
                    continue
                make_labels_now(profiles)
            elif v == "folders":
                folders_menu()
    finally:
        tui_stop()
    print()
    print("\n".join(banner()))
    print("\n  " + sty("All done.", GOOD) + "  Run the labels script whenever your spreadsheets are ready.\n")
    return 0


# ════════════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(description="Exhibition Label Studio")
    ap.add_argument("mode", nargs="?", default="generate", choices=["generate", "configure"],
                    help="generate (default) makes the labels; configure opens the set-up screen")
    ap.add_argument("--profile", help="use this profile name instead of choosing")
    ap.add_argument("--limit", type=int, help="only make the first N labels (for testing)")
    args = ap.parse_args()
    try:
        code = cmd_configure(args) if args.mode == "configure" else cmd_generate(args)
    except KeyboardInterrupt:
        tui_stop()
        print("\n  Cancelled.\n")
        code = 1
    sys.exit(code)


if __name__ == "__main__":
    main()
