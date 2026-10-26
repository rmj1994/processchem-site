# processchem.no

Static website for ProcessChem: the landing page and the documentation.
Served by GitHub Pages from `main`. No build step runs on the server.

- `index.html`, `assets/`: the landing page and shared styles, fonts and icons.
- `docs/`: **generated**. Do not edit by hand.
- `tools/build_docs.py`: builds `docs/`, `sitemap.xml`, `robots.txt` and `404.html`
  from the documentation pages (Markdown), then checks every link.

```
python3 tools/build_docs.py --source ~/Documents/GitHub/ProcessChem-Doc.wiki
```

Needs only the `markdown` Python package. Fonts are self-hosted (SIL OFL,
see `assets/fonts/LICENSES.txt`); the site loads nothing from other hosts.

App screenshots (the landing page's phones and the pictures in the docs) come
from the app itself. After a change to the app's screens:

```
python3 tools/screens.py
python3 tools/build_docs.py
```

`screens.py` starts the app's web build, captures every screen with the app's
`tools/app_shots.py`, and writes `assets/screens/*.jpg` and the wiki's
`images/screens/*.jpg`. Commit both repositories.
