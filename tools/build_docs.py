#!/usr/bin/env python3
"""Build processchem.no/docs from the ProcessChem documentation wiki.

    python3 tools/build_docs.py [--source PATH_TO_WIKI_CLONE]

Each wiki page `Heater.md` becomes `docs/Heater/index.html`, so the URL is
`/docs/Heater/` -- the same page names the app already links to on the wiki.
`ProcessChem-Documentation.md` (the wiki's Home) becomes `docs/index.html`.
Also writes `sitemap.xml`, `robots.txt` and `404.html`, then checks the result:
every internal link and #anchor must resolve, every image must exist, and no
page may load a script, stylesheet, font or image from another host.

Needs only the `markdown` package. Until the next app release the wiki is the
source of truth: edit there, rerun this, commit, push.

Colours: the shared palette lives in the app repo's
`pocketchem/design_tokens.json`. This script copies it into a generated block
at the top of `assets/site.css` (between the `app-tokens` markers), so a
colour changed in the app reaches the site on the next build. Edit the JSON,
never the block. Site-only colours (mint, the bands, the company-profile
demo colours) stay hand-written in site.css below the block.
"""

import argparse
import html
import json
import re
import shutil
import sys
from pathlib import Path

import markdown

SITE = Path(__file__).resolve().parent.parent
BASE = "https://processchem.no"
HOME_MD = "ProcessChem-Documentation"
DEFAULT_SOURCE = Path.home() / "Documents/GitHub/ProcessChem-Doc.wiki"
DEFAULT_TOKENS = Path.home() / "Documents/GitHub/PocketChem-ios/pocketchem/design_tokens.json"

#: site.css variable <- app colour token. The site's older names are kept so
#: no rule below the block has to change.
TOKEN_VARS = (
    ("--bg", "bg"), ("--surface", "surface"), ("--soft", "nav"),
    ("--ink", "ink"), ("--muted", "muted"), ("--brand", "brand"),
    ("--btn", "primary"), ("--btn-ink", "on_primary"), ("--pill", "pill"),
    # Not used by the site yet; here so new rules can match the app exactly.
    ("--card-line", "line"), ("--field-border", "field_border"),
    ("--info-bg", "info_bg"), ("--info-fg", "info_fg"),
    ("--warn-bg", "warn_bg"), ("--warn-fg", "warn_fg"),
    ("--error-bg", "error_bg"), ("--error-fg", "error_fg"),
    ("--success-bg", "success_bg"), ("--success-fg", "success_fg"),
    ("--hot", "hot"), ("--cold", "cold"),
)
TOKENS_START = "/* app-tokens:start"
TOKENS_END = "/* app-tokens:end */"

# Wording that only makes sense on the GitHub wiki itself.
HOME_FIXES = (("This wiki is its documentation", "This is its documentation"),)
HOME_DROP_SECTIONS = ("Contributing",)


def github_slug(value: str, separator: str = "-") -> str:
    """Heading id the way GitHub makes it, so wiki anchors keep working."""
    value = re.sub(r"<[^>]+>", "", value).strip().lower()
    value = re.sub(r"[^\w\- ]", "", value, flags=re.UNICODE)
    return value.replace(" ", separator)


def drop_sections(text: str, names) -> str:
    out, skipping = [], False
    for line in text.splitlines():
        m = re.match(r"^## (.+?)\s*$", line)
        if m:
            skipping = m.group(1) in names
        if not skipping:
            out.append(line)
    return "\n".join(out) + "\n"


def parse_nav(home_text: str):
    """[(group heading, [(link title, page name)])] from the home page's lists."""
    groups, current = [], None
    for line in home_text.splitlines():
        m = re.match(r"^## (.+?)\s*$", line)
        if m:
            current = (m.group(1), [])
            groups.append(current)
            continue
        m = re.match(r"^- \[([^\]]+)\]\(([^)#]+)\)", line)
        if m and current is not None:
            current[1].append((m.group(1), m.group(2)))
    return [g for g in groups if g[1]]


def render(text: str):
    md = markdown.Markdown(
        extensions=["tables", "fenced_code", "toc"],
        extension_configs={"toc": {"slugify": github_slug}},
        output_format="html5",
    )
    body = md.convert(text)
    body = re.sub(r"<table>", '<div class="table-wrap"><table>', body)
    body = body.replace("</table>", "</table></div>")
    body = SHOTS_P.sub(shots_html, body)
    return body


#: A paragraph holding nothing but app screenshots (images/screens/...): shown as a row
#: of phone pictures, each captioned with its alt text and opening full size on a tap.
SHOTS_P = re.compile(r'<p>((?:\s*<img [^>]*src="images/screens/[^"]+"[^>]*>\s*)+)</p>')
SHOT_IMG = re.compile(r'<img alt="([^"]*)" src="(images/screens/[^"]+)"\s*/?>')
SHOT_SIZE = (600, 1298)   # tools/screens.py writes the docs pictures at this size


def shots_html(m) -> str:
    figs = []
    for alt, src in SHOT_IMG.findall(m.group(1)):
        figs.append(
            f'<figure><a href="{src}"><img src="{src}" alt="{alt}" width="{SHOT_SIZE[0]}" '
            f'height="{SHOT_SIZE[1]}" loading="lazy"></a><figcaption>{alt}</figcaption></figure>')
    return '<div class="doc-shots">' + "".join(figs) + "</div>"


def page_url(name: str) -> str:
    return "/docs/" if name == HOME_MD else f"/docs/{name}/"


def rewrite_links(body: str, pages: dict, problems: list, where: str) -> str:
    lower = {n.lower(): n for n in pages}

    def href(m):
        target = m.group(1)
        if re.match(r"^(https?:|mailto:|/|#|images/)", target):
            return m.group(0)
        page, _, anchor = target.partition("#")
        hit = pages.get(page) and page or lower.get(page.lower())
        if hit is None:
            problems.append(f"{where}: link to unknown page '{page}'")
            return m.group(0)
        return 'href="%s%s"' % (page_url(hit), ("#" + anchor) if anchor else "")

    body = re.sub(r'href="([^"]*)"', href, body)
    body = re.sub(r'(src|href)="images/([^"]+)"', r'\1="/docs/images/\2"', body)
    return body


def description_of(text: str) -> str:
    """First paragraph after the title/About heading, as plain text."""
    for block in re.split(r"\n\s*\n", text):
        block = block.strip()
        if not block or block.startswith("#") or block.startswith(("-", "|", "```")):
            continue
        plain = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", block)
        plain = re.sub(r"[*_`]", "", plain)
        plain = " ".join(plain.split())
        return plain if len(plain) <= 155 else plain[:152].rsplit(" ", 1)[0] + "..."
    return "ProcessChem documentation."


HEAD = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{title}</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{canon}">
<meta name="color-scheme" content="light">
<meta name="theme-color" content="#F8FAFF">
<link rel="icon" type="image/png" href="/assets/favicon.png">
<link rel="apple-touch-icon" href="/assets/apple-touch-icon.png">
<meta property="og:type" content="article">
<meta property="og:site_name" content="ProcessChem">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:url" content="{canon}">
<meta property="og:image" content="{base}/assets/og-image.png">
<link rel="preload" href="/assets/fonts/roboto-latin-700-normal.woff2" as="font" type="font/woff2" crossorigin>
<link rel="stylesheet" href="/assets/site.css">
</head>
<body>
<a class="skip" href="#main">Skip to content</a>
<header class="site-head">
  <div class="wrap top">
    <a class="brand" href="/"><img src="/assets/icon-512.png" width="34" height="34" alt="">ProcessChem</a>
    <button class="menu-btn" type="button" aria-label="Menu" aria-expanded="false" aria-controls="nav"><span></span></button>
    <nav class="nav" id="nav" aria-label="Main">
      <a href="/#screens">The app</a>
      <a href="/#inside">What it does</a>
      <a href="/#companies">For companies</a>
      <a href="/docs/" aria-current="page">Documentation</a>
      <a class="cta-link" href="https://apps.apple.com/app/id6796077930">Get the app</a>
    </nav>
  </div>
  <div class="progress" aria-hidden="true"></div>
</header>
<div class="wrap">
"""

FOOT = """</div>
<footer class="band cut-top site-foot">
  <div class="wrap">
    <div>
      <strong>ProcessChem</strong><br>
      Built by Ron M. Jacob<br>
      <a href="mailto:ronmjacob1994@gmail.com" style="font-family:var(--mono);font-size:13px">ronmjacob1994@gmail.com</a>
    </div>
    <div class="links">
      <a href="/docs/">Documentation</a>
      <a href="/privacy/">Privacy policy</a>
    </div>
  </div>
</footer>
<script src="/assets/site.js" defer></script>
</body>
</html>
"""


#: The phone-width bar that opens the page list (site.js toggles it; wide screens hide it).
TOGGLE = (
    '<button class="docs-toggle" type="button" aria-expanded="false" aria-controls="docs-list">'
    '<svg class="ico" viewBox="0 0 20 20" aria-hidden="true"><path d="M3 5h14M3 10h14M3 15h14"/></svg>'
    '<span class="what">Contents</span><span class="cur">{title}</span>'
    '<svg class="chev" viewBox="0 0 20 20" aria-hidden="true"><path d="M5 8l5 5 5-5"/></svg></button>'
)


def nav_html(nav, current: str, page_title: str = "") -> str:
    home_cur = ' aria-current="page"' if current == HOME_MD else ""
    titles = {page: title for _, items in nav for title, page in items}
    here = "Overview" if current == HOME_MD else titles.get(current, page_title or current.replace("-", " "))
    out = ['<nav class="docs-nav" aria-label="Documentation">',
           TOGGLE.format(title=html.escape(here)),
           '<div class="docs-nav-list" id="docs-list">',
           f'<ul><li><a href="/docs/"{home_cur}>Overview</a></li></ul>']
    for heading, items in nav:
        out.append(f"<h2>{html.escape(heading)}</h2><ul>")
        for title, page in items:
            cur = ' aria-current="page"' if page == current else ""
            out.append(f'<li><a href="{page_url(page)}"{cur}>{html.escape(title)}</a></li>')
        out.append("</ul>")
    out.append("</div>")
    out.append("</nav>")
    return "\n".join(out)


def build_privacy() -> str:
    """The privacy policy, kept in privacy/PRIVACY.md and published at /privacy/."""
    src = SITE / "privacy" / "PRIVACY.md"
    text = src.read_text()
    out = SITE / "privacy" / "index.html"
    body = render(text)
    page = (
        HEAD.format(title="Privacy policy | ProcessChem", desc="ProcessChem collects no personal data. How the app and this website handle information.",
                    canon=BASE + "/privacy/", base=BASE).replace(' aria-current="page"', "")
        + '  <main id="main" class="article" style="padding-block:24px 40px">\n' + body + "\n  </main>\n" + FOOT
    )
    out.write_text(page)
    return BASE + "/privacy/"


# ---------------------------------------------------------------- landing page from content/landing.md

SLOT_RE = re.compile(r"\{\{([A-Za-z0-9_.]+)(?:\|([a-z]+))?\}\}")
SLOT_HEAD_RE = re.compile(r"^###\s+(.*?)\s*\{#([A-Za-z0-9_.]+)\}\s*$")
GENERATED_NOTE = ("<!-- Generated by tools/build_docs.py from index.template.html and "
                  "content/landing.md. Edit those files, not this one. -->\n")


def parse_landing(text: str) -> dict:
    """{slot key: text or [list items]} from the Markdown content file."""
    slots, key, buf = {}, None, []

    def close():
        if key is None:
            return
        body = [l for l in buf]
        while body and not body[0].strip():
            body.pop(0)
        while body and not body[-1].strip():
            body.pop()
        if body and all(re.match(r"^[-*+]\s+", l) or not l.strip() for l in body):
            slots[key] = [re.sub(r"^[-*+]\s+", "", l).strip() for l in body if l.strip()]
        else:
            paras = [p.strip() for p in "\n".join(body).split("\n\n") if p.strip()]
            if len(paras) != 1:
                raise ValueError(f"slot '{key}' must hold exactly one paragraph (found {len(paras)})")
            slots[key] = " ".join(paras[0].split("\n"))

    for line in text.split("\n"):
        if re.match(r"^#{1,6}\s", line):
            close()
            key, buf = None, []
            m = SLOT_HEAD_RE.match(line)
            if m:
                key = m.group(2)
            elif line.startswith("### "):
                raise ValueError(f"heading without a hidden marker: {line!r}. New text slots must "
                                 f"be added through Claude first.")
        elif key is not None:
            buf.append(line)
    close()
    return slots


def inline_html(text: str) -> str:
    """Tiny inline Markdown: **bold**, *italic*, [link](url); everything else is plain text."""
    text = re.sub(r"\\([\\`*_{}\[\]()#+\-.!|<>~&])", lambda m: "\x00%d\x00" % ord(m.group(1)), text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2">\1</a>', text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<em>\1</em>", text)
    text = re.sub("\x00(\\d+)\x00", lambda m: html.escape(chr(int(m.group(1))), quote=False), text)
    return text


def plain_text(text: str) -> str:
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"\1", text)
    return re.sub(r"\\([\\`*_{}\[\]()#+\-.!|<>~&])", r"\1", text)


def render_landing(template: str, slots: dict) -> str:
    used = set()

    def sub(m):
        key, mode = m.group(1), m.group(2) or ""
        if key not in slots:
            raise ValueError(f"template needs slot '{key}' which is missing from content/landing.md")
        used.add(key)
        val = slots[key]
        if mode == "li":
            if not isinstance(val, list):
                raise ValueError(f"slot '{key}' must be a bullet list")
            return "".join("<li>" + inline_html(v) + "</li>" for v in val)
        if isinstance(val, list):
            raise ValueError(f"slot '{key}' is a list but is used as text")
        if mode == "attr":
            return plain_text(val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        if mode == "json":
            return json.dumps(plain_text(val), ensure_ascii=False)[1:-1]
        return inline_html(val)

    out = SLOT_RE.sub(sub, template)
    unused = sorted(set(slots) - used)
    if unused:
        raise ValueError("slots in content/landing.md that the page does not use: " + ", ".join(unused))
    return out


def build_landing(check_only: bool = False) -> str:
    template = (SITE / "index.template.html").read_text(encoding="utf-8")
    slots = parse_landing((SITE / "content" / "landing.md").read_text(encoding="utf-8"))
    page = render_landing(template, slots)
    page = page.replace("<!doctype html>\n", "<!doctype html>\n" + GENERATED_NOTE, 1) \
        if GENERATED_NOTE not in page else page
    if not check_only:
        (SITE / "index.html").write_text(page, encoding="utf-8")
    return page


def build(source: Path):
    if not source.is_dir():
        sys.exit(f"source not found: {source}")
    md_files = sorted(p for p in source.glob("*.md") if not p.name.startswith("_"))
    pages = {p.stem: p for p in md_files}
    if HOME_MD not in pages:
        sys.exit(f"{HOME_MD}.md missing in {source}")

    home_text = pages[HOME_MD].read_text()
    for old, new in HOME_FIXES:
        home_text = home_text.replace(old, new)
    nav = parse_nav(home_text)
    listed = {p for _, items in nav for _, p in items}
    unlisted = sorted(set(pages) - listed - {HOME_MD})

    out_root = SITE / "docs"
    if out_root.exists():
        shutil.rmtree(out_root)
    out_root.mkdir()
    if (source / "images").is_dir():
        shutil.copytree(source / "images", out_root / "images")

    problems, built = [], []
    for name, path in pages.items():
        text = home_text if name == HOME_MD else path.read_text()
        if name == HOME_MD:
            text = drop_sections(text, HOME_DROP_SECTIONS)
        title_m = re.search(r"^# (.+)$", text, re.M)
        title = title_m.group(1).strip() if title_m else name.replace("-", " ")
        body = rewrite_links(render(text), pages, problems, name)
        canon = BASE + page_url(name)
        full_title = "ProcessChem documentation" if name == HOME_MD else f"{title} | ProcessChem documentation"
        crumb = "" if name == HOME_MD else '<div class="crumb"><a href="/docs/">Documentation</a></div>'
        page = (
            HEAD.format(title=html.escape(full_title), desc=html.escape(description_of(text)), canon=canon, base=BASE)
            + '  <div class="docs-layout">\n'
            + nav_html(nav, name, title) + "\n"
            + f'  <main id="main" class="article">\n{crumb}\n{body}\n  </main>\n  </div>\n'
            + FOOT
        )
        dest = out_root / "index.html" if name == HOME_MD else out_root / name / "index.html"
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(page)
        built.append(name)

    # sitemap, robots, 404
    build_landing()
    privacy_url = build_privacy()
    urls = [BASE + "/", privacy_url] + [BASE + page_url(n) for n in sorted(built, key=lambda n: (n != HOME_MD, n))]
    (SITE / "sitemap.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "".join(f"  <url><loc>{u}</loc></url>\n" for u in urls) + "</urlset>\n"
    )
    (SITE / "robots.txt").write_text(f"User-agent: *\nAllow: /\n\nSitemap: {BASE}/sitemap.xml\n")
    (SITE / "404.html").write_text(
        HEAD.format(title="Page not found | ProcessChem", desc="This page does not exist.", canon=BASE + "/404.html", base=BASE)
        .replace('<meta property="og:type" content="article">', '<meta name="robots" content="noindex">\n<meta property="og:type" content="website">')
        + '  <main id="main" class="article" style="padding-block:60px">\n'
        '    <p class="tag">404</p>\n    <h1>That page is not here.</h1>\n'
        '    <p>Try the <a href="/">home page</a> or the <a href="/docs/">documentation</a>.</p>\n  </main>\n'
        + FOOT
    )
    return built, unlisted, problems


def check(built, problems):
    """Every internal link/anchor resolves, images exist, nothing loads off-site."""
    html_files = [SITE / "index.html", SITE / "404.html", SITE / "privacy" / "index.html"] + sorted((SITE / "docs").rglob("index.html"))
    ids = {}

    def url_to_file(url_path: str):
        p = url_path.lstrip("/")
        if p == "" or p.endswith("/"):
            return SITE / p / "index.html"
        return SITE / p

    for f in html_files:
        ids[f] = set(re.findall(r'\bid="([^"]+)"', f.read_text()))

    for f in html_files:
        raw = f.read_text()
        text = raw
        rel = f.relative_to(SITE)
        for m in re.finditer(r'<script\b[^>]*?src="(https?://[^"]+)"', raw):
            problems.append(f"{rel}: loads from another host: {m.group(1)}")
        # markup inside <script> blocks is JavaScript strings, not links
        text = re.sub(r"<script\b[^>]*>.*?</script>", "", raw, flags=re.S)
        for href in re.findall(r'(?:href|src)="([^"]*)"', text):
            if re.match(r"^(mailto:|data:)", href):
                continue
            if re.match(r"^https?://", href):
                continue
            path, _, anchor = href.partition("#")
            if path == "":
                target = f
            else:
                target = url_to_file(path)
                if not target.exists():
                    problems.append(f"{rel}: '{href}' does not exist")
                    continue
            if anchor and target.suffix == ".html":
                if target not in ids:
                    ids[target] = set(re.findall(r'\bid="([^"]+)"', target.read_text()))
                if anchor not in ids[target]:
                    problems.append(f"{rel}: anchor '#{anchor}' not found in {target.relative_to(SITE)}")
        # resources only: nothing may be fetched from another host
        for m in re.finditer(r'<(?:script|img|iframe|source|video|audio)\b[^>]*?src="(https?://[^"]+)"', text):
            problems.append(f"{rel}: loads from another host: {m.group(1)}")
        for m in re.finditer(r'<link\b[^>]*?rel="(?:stylesheet|preload|icon)"[^>]*?href="(https?://[^"]+)"', text):
            problems.append(f"{rel}: loads from another host: {m.group(1)}")
        for host in ("googleapis", "gstatic", "cloudflare", "jsdelivr", "unpkg"):
            if host in text:
                problems.append(f"{rel}: mentions {host}")
    for css in (SITE / "assets" / "site.css",):
        t = css.read_text()
        if re.search(r"url\((?!['\"]?/assets/|['\"]?data:)", t):
            problems.append("site.css: url() outside /assets/")


def sync_tokens(tokens_path: Path, css_path: Path) -> str:
    """Rewrite site.css's generated colour block from the app's tokens.

    Returns a one-line report. A missing tokens file is reported, not fatal:
    the block already in site.css stays as it is, so the site still builds on
    a machine without the app checked out.
    """
    if not tokens_path.exists():
        return f"tokens: {tokens_path} not found -- colour block left as it is"
    colours = json.loads(tokens_path.read_text())["color"]
    lines = [f"  {var}: {colours[key]};" for var, key in TOKEN_VARS]
    block = (f"{TOKENS_START} -- generated by tools/build_docs.py from the app's "
             "pocketchem/design_tokens.json; edit the colours there, not here */\n"
             + "\n".join(lines) + f"\n  {TOKENS_END}")
    css = css_path.read_text()
    start, end = css.find(TOKENS_START), css.find(TOKENS_END)
    if start < 0 or end < 0:
        raise SystemExit("site.css has no app-tokens block to fill")
    updated = css[:start] + block + css[end + len(TOKENS_END):]
    if updated != css:
        css_path.write_text(updated)
        return f"tokens: site.css colours updated from {tokens_path.name}"
    return "tokens: site.css colours already match the app"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    ap.add_argument("--tokens", type=Path, default=DEFAULT_TOKENS)
    args = ap.parse_args()
    print(sync_tokens(args.tokens, SITE / "assets" / "site.css"))
    built, unlisted, problems = build(args.source)
    check(built, problems)
    print(f"built {len(built)} pages: {', '.join(sorted(built))}")
    if unlisted:
        print("NOTE: not in the Overview lists (still built, not in the side nav): " + ", ".join(unlisted))
    if problems:
        print(f"\n{len(problems)} PROBLEM(S):")
        for p in problems:
            print("  -", p)
        sys.exit(1)
    print("link check OK: all internal links and anchors resolve; no third-party resources")


if __name__ == "__main__":
    main()
