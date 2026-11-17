# Archived: "For companies" section

Taken off processchem.no on 2026-11-17 at the owner's request, to be added back later.
The folder name starts with `_`, so GitHub Pages (Jekyll) does not publish it.

What is here, exactly as it was on the landing page:

- `section.html`: the `<section id="companies">` markup and the inline script that drew the
  example phone (Standard app / Acme Process Co. / Northwind Chemicals). The script went
  at the end of `index.template.html`, before `/assets/site.js`.
- `landing.md`: the copy slots (`nav.companies`, `companies.*`) from `content/landing.md`.
- `site.css`: the styles for the example phone (`.app`), the demo card (`.demo`, `.seg`) and
  the three profile colours, taken out of `assets/site.css`.

To bring it back:

1. Paste `section.html`'s section between the "What it stands on" section and `#docs` in
   `index.template.html`, and its script before `<script src="/assets/site.js" defer>`.
2. Add `<a href="#companies">{{nav.companies}}</a>` to the landing page menu, and
   `<a href="/#companies">For companies</a>` to the menu in `tools/build_docs.py` (`HEAD`).
3. Paste `landing.md`'s slots back into `content/landing.md` (the build refuses slots the page
   does not use, and slots the page needs but the file lacks).
4. Append `site.css` to `assets/site.css` (the `--p-*` variables go back in `:root`).
5. Rebuild with `tools/build_docs.py` and check `link check OK`.

The Word copy of the landing page (`Website - Landing page.docx`) still holds this text until it
is next synced.
