#!/usr/bin/env python3
"""Edit ProcessChem's website, documentation and policy text in Microsoft Word.

Markdown stays the real source. This tool writes it out as .docx files for you to read and
edit in Word, and later reads your edited .docx files back, changing ONLY the paragraphs,
list items, table and code blocks that you really edited. Everything else in the source
files stays byte-for-byte as it was. Nothing is written without --apply.

    python3 tools/word_sync.py export [--only NAME] [--force]
    python3 tools/word_sync.py status
    python3 tools/word_sync.py import [--apply] [--force] [--all]
    python3 tools/word_sync.py selftest

Needs pandoc (3.x) and Python's standard library only. Config: tools/word_sources.json.
"""

import argparse
import datetime
import difflib
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
SITE = HERE.parent
CONFIG = HERE / "word_sources.json"
PANDOC = shutil.which("pandoc") or "/opt/homebrew/bin/pandoc"
SENT = "<!--@@-->"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

CURLY = "‘’“”"
STRAIGHTEN = {"‘": "'", "’": "'", "“": '"', "”": '"', " ": " ", "​": ""}


class SyncError(Exception):
    """A problem the user should read, not a crash."""


# --------------------------------------------------------------------------- config

def expand(p: str) -> Path:
    return Path(p).expanduser()


def load_config(path: Path = CONFIG, overrides: dict = None) -> dict:
    cfg = json.loads(path.read_text())
    if overrides:
        cfg.update(overrides)
    cfg["word_dir_path"] = expand(cfg["word_dir"])
    cfg["roots_path"] = {k: expand(v) for k, v in cfg["roots"].items()}
    return cfg


def source_path(cfg: dict, ref: str) -> Path:
    root, _, rel = ref.partition(":")
    if root not in cfg["roots_path"]:
        raise SyncError(f"Unknown root '{root}' in {ref}")
    return cfg["roots_path"][root] / rel


def stem_of(ref: str) -> str:
    return Path(ref.partition(":")[2]).stem


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# --------------------------------------------------------------------------- pandoc

def pandoc(args, data: str = None, binary_out: bool = False):
    cmd = [PANDOC] + list(args)
    res = subprocess.run(cmd, input=None if data is None else data.encode("utf-8"),
                         capture_output=True)
    if res.returncode != 0:
        raise SyncError("pandoc failed: " + res.stderr.decode("utf-8", "replace")[:400])
    return res.stdout if binary_out else res.stdout.decode("utf-8")


# --------------------------------------------------------------------------- icon tokens

IMG_RE = re.compile(r"!\[([^\]]*)\]\(images/([A-Za-z0-9._-]+?)\.svg\)")
TOKEN_RE = re.compile(r"\{\{icon:([A-Za-z0-9._-]+)\}\}")


def build_icon_table(texts) -> dict:
    """icon name -> alt text, checked to be the same everywhere it is used."""
    table = {}
    for text in texts:
        for alt, name in IMG_RE.findall(text):
            if table.setdefault(name, alt) != alt:
                raise SyncError(f"Icon '{name}' is used with two different alt texts; "
                                f"make them the same first.")
    return table


def tokenise(text: str) -> str:
    return IMG_RE.sub(lambda m: "{{icon:%s}}" % m.group(2), text)


def detokenise(text: str, icons: dict) -> str:
    def back(m):
        name = m.group(1)
        if name not in icons:
            raise SyncError(f"Unknown icon token {{{{icon:{name}}}}}. "
                            f"Known icons: {', '.join(sorted(icons)) or 'none'}")
        return f"![{icons[name]}](images/{name}.svg)"
    return TOKEN_RE.sub(back, text)


# --------------------------------------------------------------------------- block segmenter

class Block:
    __slots__ = ("kind", "text", "sep")

    def __init__(self, kind, text, sep="\n\n"):
        self.kind, self.text, self.sep = kind, text, sep

    def __repr__(self):
        return f"Block({self.kind!r}, {self.text[:30]!r})"


FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})")
HEADING_RE = re.compile(r"^\s{0,3}#{1,6}(\s|$)")
ITEM_RE = re.compile(r"^(\s{0,3})([-*+]|\d{1,9}[.)])(\s+|$)")
TABLE_RE = re.compile(r"^\s*\|")


def segment(text: str):
    """Split Markdown into editable blocks, remembering the exact spacing between them.

    Kinds: heading, para, item (one list item incl. nested lines), table, code.
    Covers the Markdown these documents actually use; anything else is a paragraph.
    """
    lines = text.split("\n")
    trailing = text.endswith("\n")
    if trailing:
        lines = lines[:-1]
    blocks, i, n = [], 0, len(lines)

    def blank_run(j):
        k = j
        while k < n and not lines[k].strip():
            k += 1
        return k

    while i < n:
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        fm = FENCE_RE.match(line)
        if fm:
            marker = fm.group(1)
            j = i + 1
            while j < n and not (lines[j].strip().startswith(marker[0] * len(marker))
                                 and set(lines[j].strip()) <= {marker[0]}):
                j += 1
            j = min(j + 1, n)
            kind, chunk = "code", lines[i:j]
        elif HEADING_RE.match(line):
            kind, chunk, j = "heading", [line], i + 1
        elif TABLE_RE.match(line):
            j = i
            while j < n and TABLE_RE.match(lines[j]):
                j += 1
            kind, chunk = "table", lines[i:j]
        elif ITEM_RE.match(line):
            j = i + 1
            while j < n:
                nxt = lines[j]
                if nxt.strip():
                    if re.match(r"^\s{2,}\S", nxt) or re.match(r"^\t", nxt):
                        j += 1
                        continue
                    break
                k = blank_run(j)
                if k < n and re.match(r"^\s{2,}\S", lines[k]):
                    j = k + 1  # blank line inside a loose item
                    continue
                break
            kind, chunk = "item", lines[i:j]
        else:
            j = i + 1
            while j < n and lines[j].strip() and not (
                    FENCE_RE.match(lines[j]) or HEADING_RE.match(lines[j])
                    or TABLE_RE.match(lines[j]) or ITEM_RE.match(lines[j])):
                j += 1
            kind, chunk = "para", lines[i:j]
        k = blank_run(j)
        sep = "\n" * (1 + (k - j))
        if k >= n:
            sep = "\n" if trailing else ""
        blocks.append(Block(kind, "\n".join(chunk), sep))
        i = k
    return blocks


def assemble(blocks) -> str:
    return "".join(b.text + b.sep for b in blocks)


# --------------------------------------------------------------------------- canonical form

def _straight(s: str) -> str:
    for a, b in STRAIGHTEN.items():
        s = s.replace(a, b)
    return s


def inl(nodes) -> str:
    """Canonical text of inline content, formatting compared per character.

    Word and Markdown nest bold/italic/code/links differently for the same visible result
    (for example bold text that contains inline code), so each character carries a set of
    formats, spaces carry only 'code'/'link', and runs of equal formats are merged.
    """
    chars = []   # (character, frozenset of formats)

    def put(ch, fmt):
        if "code" in fmt:
            fmt = frozenset(x for x in fmt if x == "code" or x.startswith("a:"))
        if ch.isspace() and ch != "\n":
            if chars and chars[-1][0] == " ":
                return
            ch = " "
            fmt = frozenset(x for x in fmt if x == "code" or x.startswith("a:"))
        chars.append((ch, fmt))

    def walk(ns, fmt):
        for n in ns:
            t, c = n["t"], n.get("c")
            if t == "Str":
                for ch in _straight(c):
                    put(ch, fmt)
            elif t in ("Space", "SoftBreak"):
                put(" ", fmt)
            elif t == "LineBreak":
                put("\n", frozenset())
            elif t == "Emph":
                walk(c, fmt | {"i"})
            elif t == "Strong":
                walk(c, fmt | {"b"})
            elif t in ("Underline", "Strikeout", "Superscript", "Subscript", "SmallCaps"):
                walk(c, fmt | {t})
            elif t == "Quoted":
                q = '"' if c[0]["t"] == "DoubleQuote" else "'"
                put(q, fmt)
                walk(c[1], fmt)
                put(q, fmt)
            elif t == "Code":
                for ch in _straight(c[1]):
                    put(ch, fmt | {"code"})
            elif t == "Link":
                walk(c[1], fmt | {"a:" + c[2][0]})
            elif t == "Image":
                for ch in "![" + inl(c[1]) + "](" + c[2][0] + ")":
                    put(ch, fmt)
            elif t in ("Span", "Cite"):
                walk(c[1], fmt)
            elif t == "Math":
                for ch in "$" + c[1] + "$":
                    put(ch, fmt)
            elif t == "RawInline":
                for ch in c[1]:
                    put(ch, fmt)
            elif t == "Note":
                for ch in "[^note]":
                    put(ch, fmt)
            else:
                for ch in "?" + t:
                    put(ch, fmt)

    walk(nodes, frozenset())
    while chars and chars[0][0] == " ":
        chars.pop(0)
    while chars and chars[-1][0] == " ":
        chars.pop()
    out, i = [], 0
    while i < len(chars):
        j = i
        while j < len(chars) and chars[j][1] == chars[i][1]:
            j += 1
        text = "".join(ch for ch, _ in chars[i:j])
        fmt = chars[i][1]
        out.append("⟦" + ",".join(sorted(fmt)) + "|" + text + "⟧" if fmt else text)
        i = j
    return "".join(out)


def blk(nodes) -> str:
    return "\n".join(b for b in (one(x) for x in nodes) if b)


def one(n) -> str:
    t, c = n["t"], n.get("c")
    if t in ("Para", "Plain"):
        return "P:" + inl(c)
    if t == "Header":
        return "H%d:%s" % (c[0], inl(c[2]))
    if t == "BulletList":
        return "UL[" + "|".join(blk(item) for item in c) + "]"
    if t == "OrderedList":
        return "OL[" + "|".join(blk(item) for item in c[1]) + "]"
    if t == "CodeBlock":
        body = "\n".join(_straight_code(l).rstrip() for l in c[1].split("\n")).strip("\n")
        return "CODE:" + body
    if t == "BlockQuote":
        return "Q[" + blk(c) + "]"
    if t == "HorizontalRule":
        return "HR"
    if t == "Div":
        return blk(c[1])
    if t == "RawBlock":
        return "" if c[1].strip() == SENT else "RAW:" + c[1]
    if t == "Table":
        rows = []
        head = c[3][1]
        for r in head:
            rows.append([blk(cell[4]) for cell in r[1]])
        for body in c[4]:
            for r in body[2] + body[3]:
                rows.append([blk(cell[4]) for cell in r[1]])
        return "TABLE[" + "||".join("|".join(r) for r in rows) + "]"
    return "?" + t


def _straight_code(s: str) -> str:
    return _straight(s)


def canon_groups(texts) -> list:
    """Canonical meaning of each block text, in one pandoc call."""
    if not texts:
        return []
    doc = ("\n\n" + SENT + "\n\n").join(texts)
    ast = json.loads(pandoc(["-f", "gfm", "-t", "json"], doc))
    groups, cur = [], []
    for b in ast["blocks"]:
        if b["t"] == "RawBlock" and b["c"][1].strip() == SENT:
            groups.append(cur)
            cur = []
        else:
            cur.append(b)
    groups.append(cur)
    out = [blk(g) for g in groups]
    if len(out) != len(texts):
        raise SyncError(f"Internal: expected {len(texts)} groups, got {len(out)}")
    return out


# --------------------------------------------------------------------------- polishing new text

def compact_table(text: str) -> str:
    rows = []
    for ln in text.split("\n"):
        s = ln.strip()
        if not s.startswith("|"):
            rows.append(ln)
            continue
        cells = re.split(r"(?<!\\)\|", s)[1:-1]
        cells = [c.strip() for c in cells]
        if cells and all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
            cells = [(":" if c.startswith(":") else "") + "---" + (":" if c.endswith(":") else "")
                     for c in cells]
        rows.append("| " + " | ".join(cells) + " |")
    return "\n".join(rows)


def polish(block: Block, keep_curly: bool) -> str:
    t = block.text
    t = t.replace(" ", " ").replace("​", "")
    if not keep_curly:
        for a in CURLY:
            t = t.replace(a, STRAIGHTEN[a])
    t = "\n".join(l.rstrip() for l in t.split("\n"))
    if block.kind == "code":
        t = re.sub(r"^(\s*)(`{3,}|~{3,})\s*text\s*$", r"\1\2", t, count=1, flags=re.M)
    elif block.kind == "table":
        t = compact_table(t)
    elif block.kind == "item":
        t = re.sub(r"^(\s*)(\d+[.)])\s{2,}", r"\1\2 ", t)
        t = re.sub(r"^(\s*)([-*+])\s{2,}", r"\1\2 ", t)
    return t


# --------------------------------------------------------------------------- merge

ATTR_RE = re.compile(r"\s*\{#[^}]*\}\s*$")


def _canon_text(b: Block) -> str:
    """Block text as compared: heading markers like {#hero.h1} are bookkeeping, not wording."""
    return ATTR_RE.sub("", b.text) if b.kind == "heading" else b.text


def merge(original: str, imported: str, icons: dict):
    """Original text with only the really edited blocks replaced. Returns (text, n_changed)."""
    keep_curly = any(c in original for c in CURLY)
    ob = segment(tokenise(original))
    nb = segment(imported)
    oc = canon_groups([_canon_text(b) for b in ob])
    nc = canon_groups([_canon_text(b) for b in nb])
    sm = difflib.SequenceMatcher(None, oc, nc, autojunk=False)
    out, changed = [], 0
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            out.extend(ob[i1:i2])
            continue
        changed += max(i2 - i1, j2 - j1)
        old = ob[i1:i2]
        for k, j in enumerate(range(j1, j2)):
            nbk = nb[j]
            text = polish(nbk, keep_curly)
            if k < len(old):
                sep = old[k].sep
                if nbk.kind == "heading" and old[k].kind == "heading":
                    m = ATTR_RE.search(old[k].text)
                    if m and not ATTR_RE.search(text):
                        text = text + " " + m.group(0).strip()   # keep the hidden marker
            else:
                prev = out[-1] if out else None
                if prev is None:
                    sep = "\n\n"
                else:
                    if prev.sep == "":
                        prev.sep = "\n\n"
                    if prev.kind == "item" and nbk.kind == "item":
                        sep, prev.sep = prev.sep, "\n"   # new item inherits the gap that followed the list
                    else:
                        sep = "\n\n"
                        if len(prev.sep) < 2:
                            prev.sep = "\n\n"
            out.append(Block(nbk.kind, text, sep))
    if changed == 0:
        return original, 0
    # the last block must end the file the way the original did
    ends_nl = original.endswith("\n")
    if out:
        out[-1].sep = "\n" if ends_nl else ""
    return detokenise(assemble(out), icons), changed


# --------------------------------------------------------------------------- docx reading

def docx_headings(docx: Path):
    """[(level, text, [bookmark names])] for every heading paragraph, in document order."""
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    out = []
    for p in root.iter(W + "p"):
        style = p.find(f"{W}pPr/{W}pStyle")
        val = style.get(W + "val") if style is not None else ""
        m = re.fullmatch(r"Heading(\d)", val or "")
        if not m:
            continue
        text = "".join(t.text or "" for t in p.iter(W + "t"))
        names = [b.get(W + "name") for b in p.iter(W + "bookmarkStart")]
        out.append((int(m.group(1)), text.strip(), names))
    return out


def docx_to_markdown(docx: Path):
    """(markdown, comments, problems) from a .docx, accepting Word tracked changes."""
    ast = json.loads(pandoc(["-f", "docx", "-t", "json", "--track-changes=accept", str(docx)]))
    problems = []

    def walk(x):
        if isinstance(x, dict):
            if x.get("t") == "Image":
                problems.append("a picture was pasted into the document (pictures are not supported)")
            if x.get("t") == "CodeBlock":
                x["c"][0] = ["", ["text"], []]
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(ast)
    md = pandoc(["-f", "json", "-t", "gfm", "--wrap=none"], json.dumps(ast))
    comments = []
    try:
        full = json.loads(pandoc(["-f", "docx", "-t", "json", "--track-changes=all", str(docx)]))

        def find(x):
            if isinstance(x, dict):
                if x.get("t") == "Span" and "comment-start" in x["c"][0][1]:
                    attrs = dict(x["c"][0][2])
                    comments.append((attrs.get("author", ""), inl(x["c"][1]) if x["c"][1] else ""))
                for v in x.values():
                    find(v)
            elif isinstance(x, list):
                for v in x:
                    find(v)
        find(full)
    except SyncError:
        pass
    return md, comments, sorted(set(problems))


def split_pages(md: str, docx: Path, expected_stems):
    """Imported Markdown -> {stem: text of that page}, using the hidden Word bookmarks."""
    blocks = segment(md)
    starts = [i for i, b in enumerate(blocks)
              if b.kind == "heading" and re.match(r"^#\s", b.text)]
    heads = [h for h in docx_headings(docx) if h[0] == 1]
    if len(starts) != len(heads):
        raise SyncError(f"Could not match the page headings ({len(starts)} found in the text, "
                        f"{len(heads)} in the Word file). Did a page title change style?")
    pages = {}
    for idx, s in enumerate(starts):
        e = starts[idx + 1] if idx + 1 < len(starts) else len(blocks)
        marks = [n for n in heads[idx][2] if n.startswith("page-")]
        stem = marks[0][5:] if marks else (expected_stems[idx] if len(starts) == len(expected_stems) else None)
        if stem is None:
            raise SyncError(f"The page titled '{heads[idx][1]}' has no page marker and cannot be "
                            f"placed. New pages must be added through Claude first.")
        if stem in pages:
            raise SyncError(f"The page '{stem}' appears twice in the Word file.")
        chunk = blocks[s:e]
        chunk[-1].sep = "\n"
        pages[stem] = assemble(chunk)
    return pages


# --------------------------------------------------------------------------- reference style

def ensure_reference(path: Path) -> Path:
    """A pandoc reference .docx with the site's heading colours and a readable code style."""
    if path.exists():
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(subprocess.run([PANDOC, "--print-default-data-file", "reference.docx"],
                                    capture_output=True, check=True).stdout)
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                data = _restyle(data)
            zout.writestr(item, data)
    tmp.replace(path)
    return path



# Word is strict about the order of elements inside a style (the OOXML schema is a sequence).
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
             "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
             "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
             "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
             "cnfStyle", "rPr", "sectPr", "pPrChange"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
               "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
               "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose",
               "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]
ORDERS = {"rPr": RPR_ORDER, "pPr": PPR_ORDER, "tblPr": TBLPR_ORDER, "style": STYLE_ORDER}


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def reorder_styles(root) -> None:
    for el in root.iter():
        order = ORDERS.get(_local(el.tag))
        if not order:
            continue
        kids = list(el)
        rank = {n: i for i, n in enumerate(order)}
        kids_sorted = sorted(kids, key=lambda k: rank.get(_local(k.tag), len(order)))
        if kids_sorted != kids:
            for k in kids:
                el.remove(k)
            for k in kids_sorted:
                el.append(k)


def styles_order_ok(xml: bytes) -> list:
    """Names of elements whose children are out of the order Word requires (empty = fine)."""
    bad = []
    for el in ET.fromstring(xml).iter():
        order = ORDERS.get(_local(el.tag))
        if not order:
            continue
        rank = {n: i for i, n in enumerate(order)}
        ranks = [rank.get(_local(k.tag), len(order)) for k in el]
        if ranks != sorted(ranks):
            bad.append(_local(el.tag))
    return bad


def _restyle(xml: bytes) -> bytes:
    ET.register_namespace("w", W[1:-1])
    for pfx, uri in [("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                     ("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")]:
        ET.register_namespace(pfx, uri)
    root = ET.fromstring(xml)

    def style(sid):
        for s in root.iter(W + "style"):
            if s.get(W + "styleId") == sid:
                return s
        return None

    def rpr(s):
        r = s.find(W + "rPr")
        if r is None:
            r = ET.SubElement(s, W + "rPr")
        return r

    def setc(s, tag, **attrs):
        r = rpr(s)
        for old in r.findall(W + tag):
            r.remove(old)
        el = ET.SubElement(r, W + tag)
        for k, v in attrs.items():
            el.set(W + k, v)

    blue = "0F47A0"
    for sid, size in (("Heading1", "40"), ("Heading2", "30"), ("Heading3", "26"), ("Title", "48")):
        s = style(sid)
        if s is not None:
            setc(s, "rFonts", ascii="Calibri", hAnsi="Calibri", cs="Calibri")
            setc(s, "color", val=blue)
            setc(s, "sz", val=size)
            setc(s, "b")
    for sid in ("Normal", "BodyText", "FirstParagraph", "Compact"):
        s = style(sid)
        if s is not None:
            setc(s, "rFonts", ascii="Calibri", hAnsi="Calibri", cs="Calibri")
            setc(s, "sz", val="22")
    for sid in ("SourceCode", "VerbatimChar"):
        s = style(sid)
        if s is not None:
            setc(s, "rFonts", ascii="Consolas", hAnsi="Consolas", cs="Consolas")
            setc(s, "sz", val="19")
    sc = style("SourceCode")
    if sc is not None:
        ppr = sc.find(W + "pPr")
        if ppr is None:
            ppr = ET.SubElement(sc, W + "pPr")
        for old in ppr.findall(W + "shd"):
            ppr.remove(old)
        shd = ET.SubElement(ppr, W + "shd")
        shd.set(W + "val", "clear")
        shd.set(W + "color", "auto")
        shd.set(W + "fill", "EEF0F6")
    tbl = style("Table")
    if tbl is not None:
        tp = tbl.find(W + "tblPr")
        if tp is None:
            tp = ET.SubElement(tbl, W + "tblPr")
        for old in tp.findall(W + "tblBorders"):
            tp.remove(old)
        b = ET.SubElement(tp, W + "tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = ET.SubElement(b, W + side)
            e.set(W + "val", "single")
            e.set(W + "sz", "4")
            e.set(W + "space", "0")
            e.set(W + "color", "C4C6D0")
    reorder_styles(root)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


# --------------------------------------------------------------------------- export

def manifest_path(cfg) -> Path:
    return cfg["word_dir_path"] / "_snapshots" / "manifest.json"


def load_manifest(cfg) -> dict:
    p = manifest_path(cfg)
    return json.loads(p.read_text()) if p.exists() else {}


def save_manifest(cfg, data) -> None:
    p = manifest_path(cfg)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, indent=2, sort_keys=True))


def all_icon_table(cfg) -> dict:
    texts = []
    for entry in cfg["files"]:
        for ref in entry["sources"]:
            p = source_path(cfg, ref)
            if p.exists():
                texts.append(p.read_text(encoding="utf-8"))
    return build_icon_table(texts)


def export_markdown(cfg, entry) -> str:
    parts = []
    for ref in entry["sources"]:
        p = source_path(cfg, ref)
        if not p.exists():
            raise SyncError(f"Source missing: {p}")
        text = tokenise(p.read_text(encoding="utf-8"))
        lines = text.split("\n")
        for i, ln in enumerate(lines):
            if re.match(r"^#\s", ln):
                lines[i] = ln.rstrip() + " {#page-%s}" % stem_of(ref)
                break
        else:
            raise SyncError(f"{ref} has no top-level '# Title' heading")
        parts.append("\n".join(lines).strip("\n"))
    return "\n\n".join(parts) + "\n"


def cmd_export(cfg, only=None, force=False):
    wd = cfg["word_dir_path"]
    snap = wd / "_snapshots"
    snap.mkdir(parents=True, exist_ok=True)
    ref = ensure_reference(snap / "reference.docx")
    manifest = load_manifest(cfg)
    done = []
    for entry in cfg["files"]:
        if only and only.lower() not in entry["word"].lower():
            continue
        docx = wd / entry["word"]
        rec = manifest.get(entry["word"])
        if docx.exists() and rec and sha256(docx) != rec.get("docx_sha256") and not force:
            raise SyncError(f"'{entry['word']}' has been edited since it was exported. Exporting "
                            f"again would overwrite your edits. Import it first, or use --force "
                            f"(a backup copy is kept).")
        if docx.exists() and force:
            stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(docx, snap / f"backup-{stamp}-{entry['word']}")
        md = export_markdown(cfg, entry)
        (snap / (Path(entry["word"]).stem + ".md")).write_text(md, encoding="utf-8")
        pandoc(["-f", "gfm+attributes", "-t", "docx", "--wrap=none", f"--reference-doc={ref}",
                "-o", str(docx)], md)
        manifest[entry["word"]] = {
            "exported": datetime.datetime.now().isoformat(timespec="seconds"),
            "docx_sha256": sha256(docx),
            "sources": {r: sha256(source_path(cfg, r)) for r in entry["sources"]},
        }
        done.append(entry["word"])
    save_manifest(cfg, manifest)
    write_readme(cfg)
    return done


README_MD = """# How the Word files work

These Word files are an editing view of the ProcessChem website and documentation text.
Edit them normally in Word. Then tell Claude: **"update the website and docs from my Word files"**.
Claude shows you exactly what changed before anything is written.

## What you can do

- Change wording, fix typos, add sentences, paragraphs, bullet points and table rows.
- Use bold, italic, and links. Use Track Changes and comments if you like: tracked changes are
  accepted when imported, and comments are listed for Claude but never applied.
- Edit formulas inside the grey code-style paragraphs. Check them afterwards: Word's autocorrect
  can turn quotes, dashes and three dots into other characters.

## What to leave alone

- **Page titles' hidden markers and link addresses.** You may change a link's visible text, but do
  not change where it points. Pages are tied to file names that the app links to.
- **Text like {{icon:open-in-new}}.** It stands for a small icon. Keep it exactly as it is.
- Do not paste pictures or text boxes. They are not supported and the import stops with a message.
- Do not rename or move these files.

## The landing page file

- The landing page file has grey-blue labels (headings) with the text underneath. Change the
  **text** freely. You may reword a label, but do not delete a label or add a new one.
- Each text has exactly **one paragraph**. Lists of small tags are bullet points: add or remove bullets.
- The first lines (page title, search description, link preview) are what Google and LinkedIn show.

## New pages or new text slots

Ask Claude to add them first, so the website and documentation know where they belong.

## If something goes wrong

Nothing is changed without your approval. Backups of every overwritten Word file are in
`_snapshots`, and every import writes a report in `_reports`.
"""


def write_readme(cfg):
    out = cfg["word_dir_path"] / "README - how this works.docx"
    ref = cfg["word_dir_path"] / "_snapshots" / "reference.docx"
    pandoc(["-f", "gfm", "-t", "docx", "--wrap=none", f"--reference-doc={ref}", "-o", str(out)],
           README_MD)


# --------------------------------------------------------------------------- status

def cmd_status(cfg):
    manifest = load_manifest(cfg)
    wd = cfg["word_dir_path"]
    rows = []
    for entry in cfg["files"]:
        docx = wd / entry["word"]
        rec = manifest.get(entry["word"])
        if not docx.exists() or not rec:
            rows.append((entry["word"], "not exported yet", []))
            continue
        edited = sha256(docx) != rec["docx_sha256"]
        stale = [r for r in entry["sources"]
                 if not source_path(cfg, r).exists() or sha256(source_path(cfg, r)) != rec["sources"].get(r)]
        state = ("EDITED in Word" if edited else "unchanged in Word")
        if stale:
            state += ", SOURCE CHANGED since export"
        rows.append((entry["word"], state, stale))
    return rows


# --------------------------------------------------------------------------- checks

def check_links(cfg, new_texts: dict):
    """Wiki-style links must point at a page that exists, and an anchor that exists."""
    problems = []
    wiki = cfg["roots_path"]["wiki"]
    stems = {p.stem: p for p in wiki.glob("*.md")} if wiki.exists() else {}

    def slug(h):
        h = re.sub(r"\{\{icon:[^}]*\}\}|!\[[^\]]*\]\([^)]*\)", "", h).strip().lower()
        h = re.sub(r"[`*_]", "", h)
        h = re.sub(r"[^\w\- ]", "", h, flags=re.U)
        return h.replace(" ", "-")

    def anchors(stem):
        text = new_texts.get(stem)
        if text is None and stem in stems:
            text = stems[stem].read_text(encoding="utf-8")
        found = set()
        for ln in (text or "").split("\n"):
            m = re.match(r"^#{1,6}\s+(.*?)\s*$", ln)
            if m:
                found.add(slug(m.group(1)))
        return found

    for stem, text in new_texts.items():
        for target in re.findall(r"\]\(([^)\s]+)\)", text):
            if re.match(r"^(https?:|mailto:|/|#)", target) or target.startswith("images/"):
                continue
            page, _, anchor = target.partition("#")
            if page not in stems and page not in new_texts:
                problems.append(f"{stem}: link to unknown page '{page}'")
            elif anchor and anchor not in anchors(page):
                problems.append(f"{stem}: link to '#{anchor}' which does not exist in {page}")
    return problems


def check_landing(cfg, text: str):
    """The landing text must still fit the page template (every slot present, one paragraph each)."""
    root = cfg["roots_path"]["site"]
    template = root / "index.template.html"
    if not template.exists():
        return []
    sys.path.insert(0, str(HERE))
    import build_docs
    try:
        build_docs.render_landing(template.read_text(encoding="utf-8"), build_docs.parse_landing(text))
    except ValueError as err:
        return [f"landing: {err}"]
    return []


def check_play_limits(text: str):
    """Play Console limits written as '(max N)' next to a value or code block must still hold."""
    problems = []
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        m = re.search(r"\(max (\d+)\)", ln)
        if not m:
            continue
        limit, value = int(m.group(1)), None
        if ln.lstrip().startswith("|"):
            q = re.search(r"`([^`]+)`", ln)
            value = q.group(1) if q else None
        else:
            j = i + 1
            while j < len(lines) and not lines[j].startswith("```"):
                if re.match(r"^#{1,6}\s", lines[j]):
                    break
                j += 1
            if j < len(lines) and lines[j].startswith("```"):
                k = j + 1
                body = []
                while k < len(lines) and not lines[k].startswith("```"):
                    body.append(lines[k])
                    k += 1
                value = "\n".join(body).strip("\n")
        if value is not None and len(value) > limit:
            what = re.sub(r"[#|`*]", "", ln).strip()[:50]
            problems.append(f"play-listing: '{what}' is {len(value)} characters; the limit is {limit}")
    return problems


def check_tokens(stem, text, icons):
    return [f"{stem}: unknown icon token {{{{icon:{n}}}}}" for n in TOKEN_RE.findall(text)
            if n not in icons]


# --------------------------------------------------------------------------- import

def cmd_import(cfg, apply=False, force=False, process_all=False, quiet=False, rebuild=False):
    manifest = load_manifest(cfg)
    wd = cfg["word_dir_path"]
    icons = all_icon_table(cfg)
    results, problems_all, notes = [], [], []
    pending = {}   # path -> new text
    processed = set()
    for entry in cfg["files"]:
        docx = wd / entry["word"]
        rec = manifest.get(entry["word"])
        if not docx.exists():
            continue
        if rec and not process_all and sha256(docx) == rec["docx_sha256"]:
            continue
        stale = [r for r in entry["sources"] if rec and
                 sha256(source_path(cfg, r)) != rec["sources"].get(r)]
        if stale and not force:
            problems_all.append(f"{entry['word']}: the source file(s) {', '.join(stale)} changed "
                                f"after this Word file was exported, so importing could overwrite "
                                f"newer text. Re-export, or use --force after checking.")
            continue
        md, comments, bad = docx_to_markdown(docx)
        if bad:
            problems_all.append(f"{entry['word']}: " + "; ".join(bad))
            continue
        for who, text in comments:
            notes.append(f"{entry['word']}: comment by {who or 'someone'}: {text}")
        stems = [stem_of(r) for r in entry["sources"]]
        processed.add(entry["word"])
        try:
            pages = split_pages(md, docx, stems)
        except SyncError as e:
            problems_all.append(f"{entry['word']}: {e}")
            continue
        missing = [s for s in stems if s not in pages]
        extra = [s for s in pages if s not in stems]
        if missing:
            problems_all.append(f"{entry['word']}: page(s) missing from the Word file: "
                                f"{', '.join(missing)} (nothing from this file was imported)")
            continue
        if extra:
            problems_all.append(f"{entry['word']}: unknown page(s) {', '.join(extra)}; new pages "
                                f"must be added through Claude first")
            continue
        for ref in entry["sources"]:
            stem = stem_of(ref)
            path = source_path(cfg, ref)
            original = path.read_text(encoding="utf-8")
            try:
                new, n = merge(original, pages[stem], icons)
            except SyncError as err:
                problems_all.append(f"{stem}: {err}")
                continue
            if n:
                pending[path] = (ref, stem, original, new, n)
    # validation across everything that would change
    texts = {stem: new for (_, stem, _, new, _) in pending.values()}
    for stem, text in texts.items():
        problems_all.extend(check_tokens(stem, text, icons))
        if "<table" in text:
            problems_all.append(f"{stem}: a table became HTML; keep tables simple (one paragraph per cell)")
    problems_all.extend(check_links(cfg, texts))
    if "play-listing" in texts:
        problems_all.extend(check_play_limits(texts["play-listing"]))
    if "landing" in texts:
        problems_all.extend(check_landing(cfg, texts["landing"]))
    diff_text = []
    for path, (ref, stem, original, new, n) in pending.items():
        diff_text.append("".join(difflib.unified_diff(
            original.splitlines(True), new.splitlines(True), f"a/{path.name}", f"b/{path.name}", n=1)))
        results.append((ref, n))
    report = build_report(results, problems_all, notes, diff_text, apply)
    save_report(cfg, report)
    if apply and not problems_all:
        for path, (ref, stem, original, new, n) in pending.items():
            path.write_text(new, encoding="utf-8")
        for entry in cfg["files"]:
            docx = wd / entry["word"]
            if docx.exists() and entry["word"] in manifest and entry["word"] in processed:
                manifest[entry["word"]]["docx_sha256"] = sha256(docx)
                manifest[entry["word"]]["sources"] = {
                    r: sha256(source_path(cfg, r)) for r in entry["sources"]}
        save_manifest(cfg, manifest)
        if rebuild and pending:
            res = subprocess.run([sys.executable, "-I", str(HERE / "build_docs.py"), "--source",
                                  str(cfg["roots_path"]["wiki"])], capture_output=True, text=True, cwd=SITE)
            report += "\n## Website rebuild\n\n```\n" + (res.stdout + res.stderr).strip() + "\n```\n"
            if res.returncode != 0:
                problems_all.append("the website rebuild reported a problem (see above); the Markdown "
                                    "files were written, please review before committing")
    if not quiet:
        print(report)
    return {"changed": results, "problems": problems_all, "notes": notes,
            "pending": {str(p): v[3] for p, v in pending.items()}}


def build_report(results, problems, notes, diffs, applied):
    lines = ["# Import report", ""]
    if problems:
        lines += ["## NOT IMPORTED. Please read:", ""] + [f"- {p}" for p in problems] + [""]
    if results:
        lines += ["## Changes found" + (" (written)" if applied and not problems else " (dry run, nothing written)"), ""]
        lines += [f"- {ref}: {n} block(s) changed" for ref, n in results] + [""]
    else:
        lines += ["No text changes found in the Word files.", ""]
    if notes:
        lines += ["## Comments in your Word files (notes for Claude, not applied)", ""] + \
                 [f"- {n}" for n in notes] + [""]
    if diffs:
        lines += ["## Exact changes", "", "```diff"] + ["".join(diffs).rstrip("\n")] + ["```", ""]
    return "\n".join(lines)


def save_report(cfg, report):
    d = cfg["word_dir_path"] / "_reports"
    d.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    (d / f"import-{stamp}.md").write_text(report, encoding="utf-8")


# --------------------------------------------------------------------------- self-test

def _edit_docx_text(docx: Path, old: str, new: str) -> None:
    """Change one run of text inside a .docx the way a person typing in Word would."""
    tmp = docx.with_suffix(".edit")
    hit = 0
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                s = data.decode("utf-8")
                hit = s.count(old)
                s = s.replace(old, new, 1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(docx)
    if not hit:
        raise SyncError(f"selftest: phrase not found in {docx.name}: {old!r}")


def _diff_counts(orig: str, new: str):
    d = [l for l in difflib.unified_diff(orig.splitlines(), new.splitlines(), lineterm="", n=0)
         if l[:1] in "+-" and not l.startswith(("+++", "---"))]
    return sum(1 for l in d if l[0] == "+"), sum(1 for l in d if l[0] == "-"), d


def _tracked_insert(docx: Path, after_phrase: str, text: str) -> None:
    tmp = docx.with_suffix(".edit")
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                s = data.decode("utf-8")
                idx = s.index(after_phrase)
                end = s.index("</w:r>", idx) + len("</w:r>")
                ins = ('<w:ins w:id="9001" w:author="Tester" w:date="2026-10-14T10:00:00Z">'
                       '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:ins>' % text)
                data = (s[:end] + ins + s[end:]).encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(docx)


def cmd_selftest():
    cfg0 = load_config()
    tmp = Path(tempfile.mkdtemp(prefix="wordsync-"))
    results = []

    def check(name, ok, detail=""):
        results.append(ok)
        print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""))

    try:
        roots = {n: str(tmp / "roots" / n) for n in cfg0["roots"]}
        for entry in cfg0["files"]:
            for ref in entry["sources"]:
                root, _, rel = ref.partition(":")
                dst = Path(roots[root]) / rel
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_path(cfg0, ref), dst)
        shutil.copy2(SITE / "index.template.html", Path(roots["site"]) / "index.template.html")
        cfg = load_config(overrides={"word_dir": str(tmp / "word"), "roots": roots})
        wd = cfg["word_dir_path"]
        originals = {str(source_path(cfg, r)): source_path(cfg, r).read_text(encoding="utf-8")
                     for e in cfg["files"] for r in e["sources"]}
        by_name = {Path(p).name: (p, o) for p, o in originals.items()}

        def fresh():
            cmd_export(cfg, force=True)

        def run():
            return cmd_import(cfg, process_all=True, quiet=True)

        def from_markdown(word, transform, extra=()):
            """Replace a Word file with one built from edited Markdown (what a Word edit looks like)."""
            entry = next(e for e in cfg["files"] if e["word"] == word)
            md = transform(export_markdown(cfg, entry))
            pandoc(["-f", "gfm+attributes", "-t", "docx", "--wrap=none",
                    f"--reference-doc={wd / '_snapshots' / 'reference.docx'}", *extra,
                    "-o", str(wd / word)], md)

        def edited(r):
            return {Path(p).name: new for p, new in r["pending"].items()}

        fresh()
        with zipfile.ZipFile(wd / "_snapshots" / "reference.docx") as z:
            bad = styles_order_ok(z.read("word/styles.xml"))
        check("Word styles follow the element order Word requires", not bad, str(bad))
        r = run()
        check("round trip with no edits changes nothing", not r["changed"] and not r["problems"],
              json.dumps(r["changed"]) + " | " + "; ".join(r["problems"]))

        # edits made directly in the .docx, the way Word stores them
        _edit_docx_text(wd / "Docs - Cases.docx", "Heats or cools a batch", "Heats or cools a BATCH")
        _edit_docx_text(wd / "Docs - Tabs and overview.docx", "bara, barg, kPa, MPa, psi, atm",
                        "bara, barg, kPa, MPa, psi, atm, mbar")
        r = run()
        e = edited(r)
        check("typed edits in Word change only the two edited files",
              sorted(e) == ["Heater.md", "Units-Tab.md"] and not r["problems"], f"{sorted(e)} {r['problems']}")
        for name, new in e.items():
            a, rm, d = _diff_counts(by_name[name][1], new)
            check(f"{name}: exactly one line replaced", (a, rm) == (1, 1), str(d[:4]))
        fresh()

        # tracked change made in Word is accepted
        _tracked_insert(wd / "Docs - Cases.docx", "Heats or cools a batch", " INSERTED")
        r = run()
        e = edited(r)
        check("a Word tracked insertion is accepted",
              "Heater.md" in e and " INSERTED" in e["Heater.md"] and not r["problems"], str(r["problems"]))
        fresh()

        # structural edits
        def add_bullet(md):
            return re.sub(r"^(- \[Units\]\(Units-Tab\).*)$", r"\1\n- [Extra entry](Units-Tab) a new bullet",
                          md, count=1, flags=re.M)
        from_markdown("Docs - Tabs and overview.docx", add_bullet)
        r = run(); e = edited(r)
        a, rm, d = _diff_counts(by_name["ProcessChem-Documentation.md"][1], e.get("ProcessChem-Documentation.md", ""))
        check("a new bullet adds exactly one line", (a, rm) == (1, 0) and not r["problems"], f"{(a, rm)} {r['problems']} {d[:3]}")
        fresh()

        from_markdown("Docs - Tabs and overview.docx",
                      lambda md: md.replace("| Mass | kg, g, lb, tonne |", "| Mass | kg, g, lb, tonne |\n| Test row | X, Y |", 1))
        r = run(); e = edited(r)
        a, rm, d = _diff_counts(by_name["Units-Tab.md"][1], e.get("Units-Tab.md", ""))
        check("a new table row keeps the table's style", a >= 1 and "| Test row | X, Y |" in e.get("Units-Tab.md", "")
              and "| --- | --- |" in e.get("Units-Tab.md", "") and not r["problems"], f"{(a, rm)} {d[:3]}")
        fresh()

        from_markdown("Docs - Cases.docx", lambda md: md.replace("Σ L_tr", "Σ Ltr", 1))
        r = run(); e = edited(r)
        h = e.get("Heater.md", "")
        a, rm, d = _diff_counts(by_name["Heater.md"][1], h)
        check("a formula edit changes only that line and keeps the code fence",
              (a, rm) == (1, 1) and "Σ Ltr" in h and h.count("```") == by_name["Heater.md"][1].count("```"),
              f"{(a, rm)} {d[:3]}")
        fresh()

        from_markdown("Docs - Tabs and overview.docx",
                      lambda md: md.replace("**[Reaction](Reaction)**", "**[Chemical Reaction](Reaction)**", 1))
        r = run(); e = edited(r)
        t_ = e.get("Thermo-Tab.md", "")
        check("editing link text keeps the link address", "(Reaction)" in t_ and "Chemical Reaction" in t_ and not r["problems"],
              f"{r['problems']}")
        fresh()


        # landing page copy (slots are headings with hidden markers; text lives under them)
        _edit_docx_text(wd / "Website - Landing page.docx", "Process calculations you can run on the plant floor.",
                        "Process calculations for the plant floor.")
        r = run(); e = edited(r)
        new = e.get("landing.md", "")
        a, rm, d = _diff_counts(by_name["landing.md"][1], new)
        check("landing: a headline edit changes exactly one line", (a, rm) == (1, 1) and not r["problems"], f"{(a, rm)} {r['problems']}")
        sys.path.insert(0, str(HERE))
        import build_docs
        page = build_docs.render_landing((Path(roots["site"]) / "index.template.html").read_text(encoding="utf-8"),
                                         build_docs.parse_landing(new))
        check("landing: the new headline reaches the rendered page", "Process calculations for the plant floor." in page)
        fresh()

        from_markdown("Website - Landing page.docx",
                      lambda md: md.replace("### Headline {#hero.h1}", "### Main headline {#hero.h1}", 1))
        r = run(); e = edited(r)
        check("landing: renaming a label keeps its hidden marker",
              "### Main headline {#hero.h1}" in e.get("landing.md", ""), str(r["problems"]))
        fresh()

        from_markdown("Website - Landing page.docx",
                      lambda md: md.replace("- surface tension\n", "- surface tension\n- boiling point\n", 1))
        r = run(); e = edited(r)
        a, rm, d = _diff_counts(by_name["landing.md"][1], e.get("landing.md", ""))
        page = build_docs.render_landing((Path(roots["site"]) / "index.template.html").read_text(encoding="utf-8"),
                                         build_docs.parse_landing(e.get("landing.md", "")))
        check("landing: a new tag in a list adds one line and one tag on the page",
              (a, rm) == (1, 0) and "<li>boiling point</li>" in page and not r["problems"], f"{(a, rm)} {r['problems']}")
        fresh()

        from_markdown("Website - Landing page.docx",
                      lambda md: md.replace("Process calculations you can run on the plant floor.",
                                            "Process calculations you can run on the plant floor.\n\nA second paragraph.", 1))
        r = run()
        check("guard: two paragraphs in one text slot are refused", any("exactly one paragraph" in p for p in r["problems"]), str(r["problems"]))
        fresh()

        from_markdown("Website - Landing page.docx",
                      lambda md: md.replace("## Menu", "### A brand new slot\n\nSome text.\n\n## Menu", 1))
        r = run()
        check("guard: a new slot without a marker is refused", any("hidden marker" in p for p in r["problems"]), str(r["problems"]))
        fresh()


        long_value = "x" * 90
        from_markdown("Store - Google Play listing.docx",
                      lambda md: md.replace("`Thermodynamics, reactions and process calculators for engineers. Works offline.`",
                                            "`" + long_value + "`", 1))
        r = run()
        check("guard: a Play short description over 80 characters is refused",
              any("limit is 80" in p for p in r["problems"]), str(r["problems"]))
        fresh()

        # apply writes, records the new state, and a second import finds nothing
        _edit_docx_text(wd / "Docs - Cases.docx", "Heats or cools a batch", "Heats or cools a BATCH")
        r = cmd_import(cfg, apply=True, process_all=True, quiet=True)
        written = source_path(cfg, "wiki:Heater.md").read_text(encoding="utf-8")
        check("apply writes the edited paragraph", "BATCH" in written and not r["problems"], str(r["problems"]))
        r2 = cmd_import(cfg, quiet=True)
        check("after apply there is nothing left to import", not r2["changed"] and not r2["problems"], str(r2["problems"]))
        # restore pristine sources for the guard tests
        for p, o in originals.items():
            Path(p).write_text(o, encoding="utf-8")
        fresh()

        # guards: each bad edit must be refused with a message
        from_markdown("Docs - Cases.docx", lambda md: md.replace("(Vessel-Heater)", "(Vessel-Heaters)"))
        r = run()
        check("guard: a link to a page that does not exist is refused",
              any("unknown page" in p for p in r["problems"]), str(r["problems"]))
        fresh()
        from_markdown("Docs - Cases.docx", lambda md: md.replace("# Heater {#page-Heater}\n", "", 1))
        r = run()
        check("guard: a deleted page heading is refused", bool(r["problems"]), str(r["problems"]))
        fresh()
        from_markdown("Docs - Tabs and overview.docx", lambda md: md.replace("{{icon:open-in-new}}", "{{icon:nope}}", 1))
        r = run()
        check("guard: a broken icon token is refused", any("icon" in p for p in r["problems"]), str(r["problems"]))
        fresh()
        png = tmp / "pic.png"
        png.write_bytes(bytes.fromhex("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
                                      "890000000d49444154789c6360000002000001e221bc330000000049454e44ae426082"))
        from_markdown("Docs - Cases.docx", lambda md: md.replace("\n\n", "\n\n![x](pic.png)\n\n", 1),
                      extra=(f"--resource-path={tmp}",))
        r = run()
        check("guard: a pasted picture is refused", any("picture" in p for p in r["problems"]), str(r["problems"]))
        fresh()
        sp = source_path(cfg, "wiki:Heater.md")
        sp.write_text(sp.read_text(encoding="utf-8") + "\nExtra line added elsewhere.\n", encoding="utf-8")
        _edit_docx_text(wd / "Docs - Cases.docx", "Heats or cools a batch", "Heats or cools a BATCH")
        r = run()
        check("guard: a source changed after export is refused", any("changed after" in p for p in r["problems"]), str(r["problems"]))
        # export must not overwrite edits
        try:
            cmd_export(cfg)
            check("guard: export refuses to overwrite an edited Word file", False, "no error")
        except SyncError:
            check("guard: export refuses to overwrite an edited Word file", True)
        return all(results)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------------------- CLI

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("export")
    e.add_argument("--only")
    e.add_argument("--force", action="store_true")
    sub.add_parser("status")
    i = sub.add_parser("import")
    i.add_argument("--apply", action="store_true")
    i.add_argument("--force", action="store_true")
    i.add_argument("--all", action="store_true", help="also read Word files that look unchanged")
    sub.add_parser("selftest")
    args = ap.parse_args(argv)
    try:
        if args.cmd == "selftest":
            sys.exit(0 if cmd_selftest() else 1)
        cfg = load_config()
        if args.cmd == "export":
            done = cmd_export(cfg, args.only, args.force)
            print("Exported:\n  " + "\n  ".join(done))
        elif args.cmd == "status":
            for name, state, stale in cmd_status(cfg):
                print(f"{name:40s} {state}")
                for s in stale:
                    print(f"    changed source: {s}")
        elif args.cmd == "import":
            res = cmd_import(cfg, apply=args.apply, force=args.force, process_all=args.all, rebuild=True)
            sys.exit(1 if res["problems"] else 0)
    except SyncError as err:
        print("Stopped: " + str(err))
        sys.exit(2)


if __name__ == "__main__":
    main()
